<#
.SYNOPSIS
    Step 1: read-only inventory of one or all Schuler tenants.

.DESCRIPTION
    Per tenant, each workload signs in interactively and writes one JSON per area:

        tenants.json ──► Export-Inventory.ps1
                           ├─ Graph    (Entra, CA, Intune, SharePoint, Secure Score, Copilot licenses)
                           ├─ Exchange (EXO + Defender for Office 365)
                           ├─ Purview  (DLP, labels, retention)
                           ├─ Teams    (policies, external and guest access)
                           └─ Dns      (MX, SPF, DKIM, DMARC per domain)
                                  │
                                  ▼
                         exports/<tenant>/<area>.json
                         exports/<tenant>/_errors.json     failed areas, 403 = missing permission
                         exports/<tenant>/_workloads.json  exported / not licensed

    Graph runs first: its license export decides whether Exchange and Teams are
    licensed. Tenants on Lotus Notes get "not licensed" instead of errors.

.EXAMPLE
    ./Export-Inventory.ps1 -Tenant SyTech
    ./Export-Inventory.ps1 -Tenant SyTech -Workload Teams
    ./Export-Inventory.ps1            # all tenants, all workloads
#>
param(
    [string] $Tenant = '*',

    # Comma list. "pwsh ./x.ps1 -Workload A,B" passes one string, split below.
    [string[]] $Workload = @('Graph', 'Exchange', 'Purview', 'Teams', 'Dns'),

    # Wam: Windows broker, preferred. Browser: system browser for hosts where WAM
    # fails (e.g. embedded terminals). Graph always uses WAM.
    [ValidateSet('Wam', 'Browser')]
    [string] $Broker = 'Wam'
)

$ErrorActionPreference = 'Stop'

# Run order matters: Graph provides licenses and domains for the rest.
$WORKLOADS = @('Graph', 'Exchange', 'Purview', 'Teams', 'Dns')

$Workload = $Workload -split ',' | ForEach-Object Trim | Where-Object { $_ }
$unknown  = $Workload | Where-Object { $_ -notin $WORKLOADS }
if ($unknown) { throw "Unknown workload '$($unknown -join ', ')'. Allowed: $($WORKLOADS -join ', ')." }

# Module cmdlets (EXO, Teams) do not inherit $ErrorActionPreference. Forces
# failures into the catch block of Save-Area.
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

$GRAPH          = 'https://graph.microsoft.com'
$LOGIN          = 'https://login.microsoftonline.com'
$EXPORT_ROOT    = Join-Path $PSScriptRoot 'exports'
$TENANT_FILE    = Join-Path $PSScriptRoot 'tenants.json'
$HTTP_FORBIDDEN = 403
$JSON_DEPTH     = 8
$EXO_PLAN       = '^EXCHANGE_S_(STANDARD|ENTERPRISE|DESKLESS|ESSENTIALS)$'   # not FOUNDATION, it ships with Intune/EMS
$TEAMS_PLAN     = '^TEAMS1$'
$NOT_LICENSED   = 'not licensed'
$EXPORTED       = 'exported'

# Graph scopes are not listed here: the app registration in Entra holds them as
# delegated permissions (admin consent per tenant). Effective rights = consented
# permissions ∩ signed-in user's role (Global Reader). See README.


# ─── Graph areas ─────────────────────────────────────────────────────────────
# Area name ──► Graph path. Area name becomes the file name.
# Personal data is kept minimal: no names or UPNs for users.
$GraphAreas = [ordered]@{
    'org.organization'        = '/v1.0/organization'
    'org.domains'             = '/v1.0/domains'
    'org.subscriptions'       = '/v1.0/directory/subscriptions'
    'lic.subscribedSkus'      = '/v1.0/subscribedSkus'

    'entra.users'             = '/v1.0/users?$select=id,userType,accountEnabled,onPremisesSyncEnabled,assignedLicenses,createdDateTime&$top=999'
    'entra.groups'            = '/v1.0/groups?$select=id,displayName,groupTypes,securityEnabled,mailEnabled,visibility,resourceProvisioningOptions,membershipRule,onPremisesSyncEnabled&$top=999'
    'entra.applications'      = '/v1.0/applications?$select=id,appId,displayName,signInAudience,createdDateTime,passwordCredentials,keyCredentials,requiredResourceAccess'
    'entra.servicePrincipals' = '/v1.0/servicePrincipals?$select=id,appId,displayName,servicePrincipalType,appOwnerOrganizationId,accountEnabled&$top=999'
    'entra.oauthGrants'       = '/v1.0/oauth2PermissionGrants'
    'entra.authorization'     = '/v1.0/policies/authorizationPolicy'
    'entra.securityDefaults'  = '/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
    'entra.authMethods'       = '/v1.0/policies/authenticationMethodsPolicy'
    'entra.authStrengths'     = '/v1.0/policies/authenticationStrengthPolicies'
    'entra.adminConsent'      = '/v1.0/policies/adminConsentRequestPolicy'
    'entra.deviceRegistration'= '/beta/policies/deviceRegistrationPolicy'
    'entra.crossTenant'       = '/v1.0/policies/crossTenantAccessPolicy/default'
    'entra.crossTenantPartners' = '/v1.0/policies/crossTenantAccessPolicy/partners'
    'entra.groupSettings'     = '/v1.0/groupSettings'
    'entra.roleDefinitions'   = '/v1.0/roleManagement/directory/roleDefinitions'
    'entra.roleAssignments'   = '/v1.0/roleManagement/directory/roleAssignments?$expand=principal($select=id,displayName,userPrincipalName)'
    'entra.pimEligible'       = '/v1.0/roleManagement/directory/roleEligibilitySchedules'
    'entra.mfaRegistration'   = '/v1.0/reports/authenticationMethods/userRegistrationDetails?$select=id,userType,isAdmin,isMfaRegistered,isMfaCapable,isPasswordlessCapable,methodsRegistered'

    'ca.policies'             = '/v1.0/identity/conditionalAccess/policies'
    'ca.namedLocations'       = '/v1.0/identity/conditionalAccess/namedLocations'

    'sec.secureScore'         = '/v1.0/security/secureScores?$top=1'
    'sec.secureScoreProfiles' = '/v1.0/security/secureScoreControlProfiles'

    'intune.compliance'       = '/v1.0/deviceManagement/deviceCompliancePolicies?$expand=assignments'
    'intune.configurations'   = '/v1.0/deviceManagement/deviceConfigurations?$expand=assignments'
    'intune.settingsCatalog'  = '/beta/deviceManagement/configurationPolicies?$expand=assignments'
    'intune.groupPolicy'      = '/beta/deviceManagement/groupPolicyConfigurations?$expand=assignments'
    'intune.endpointSecurity' = '/beta/deviceManagement/intents'
    'intune.enrollment'       = '/v1.0/deviceManagement/deviceEnrollmentConfigurations'
    'intune.appProtection'    = '/v1.0/deviceAppManagement/managedAppPolicies'
    'intune.apps'             = '/v1.0/deviceAppManagement/mobileApps?$expand=assignments'
    'intune.autopilotProfiles'= '/beta/deviceManagement/windowsAutopilotDeploymentProfiles'
    'intune.autopilotDevices' = '/v1.0/deviceManagement/windowsAutopilotDeviceIdentities'
    'intune.featureUpdates'   = '/beta/deviceManagement/windowsFeatureUpdateProfiles'
    'intune.qualityUpdates'   = '/beta/deviceManagement/windowsQualityUpdateProfiles'
    'intune.scripts'          = '/beta/deviceManagement/deviceManagementScripts'
    'intune.remediations'     = '/beta/deviceManagement/deviceHealthScripts'
    'intune.devices'          = '/v1.0/deviceManagement/managedDevices?$select=id,operatingSystem,osVersion,complianceState,managementAgent,managedDeviceOwnerType'

    'spo.settings'            = '/v1.0/admin/sharepoint/settings'
}


# ─── Exchange Online + Defender for Office 365 ──────────────────────────────
$ExchangeAreas = [ordered]@{
    'exo.organization'      = { Get-OrganizationConfig }
    'exo.transport'         = { Get-TransportConfig }
    'exo.acceptedDomains'   = { Get-AcceptedDomain }
    'exo.remoteDomains'     = { Get-RemoteDomain }
    'exo.transportRules'    = { Get-TransportRule }
    'exo.inboundConnectors' = { Get-InboundConnector }
    'exo.outboundConnectors'= { Get-OutboundConnector }
    'exo.sharingPolicies'   = { Get-SharingPolicy }
    'exo.externalTagging'   = { Get-ExternalInOutlook }
    'exo.owaPolicies'       = { Get-OwaMailboxPolicy }
    'exo.mobilePolicies'    = { Get-MobileDeviceMailboxPolicy }
    'exo.authPolicies'      = { Get-AuthenticationPolicy }
    'exo.casPlans'          = { Get-CASMailboxPlan }
    'exo.dkim'              = { Get-DkimSigningConfig }
    'exo.auditConfig'       = { Get-AdminAuditLogConfig }
    'exo.mailboxCounts'     = { Get-EXOMailbox -ResultSize Unlimited | Group-Object RecipientTypeDetails | Select-Object Name, Count }

    'mdo.antiSpam'          = { Get-HostedContentFilterPolicy }
    'mdo.antiSpamRules'     = { Get-HostedContentFilterRule }
    'mdo.outboundSpam'      = { Get-HostedOutboundSpamFilterPolicy }
    'mdo.connectionFilter'  = { Get-HostedConnectionFilterPolicy }
    'mdo.antiMalware'       = { Get-MalwareFilterPolicy }
    'mdo.antiPhish'         = { Get-AntiPhishPolicy }
    'mdo.quarantine'        = { Get-QuarantinePolicy }
    'mdo.presetEop'         = { Get-EOPProtectionPolicyRule }
    'mdo.presetAtp'         = { Get-ATPProtectionPolicyRule }
    'mdo.safeLinks'         = { Get-SafeLinksPolicy }
    'mdo.safeAttachments'   = { Get-SafeAttachmentPolicy }
    'mdo.atpGlobal'         = { Get-AtpPolicyForO365 }
}


# ─── Purview (Security & Compliance PowerShell) ─────────────────────────────
$PurviewAreas = [ordered]@{
    'pv.dlpPolicies'       = { Get-DlpCompliancePolicy }
    'pv.dlpRules'          = { Get-DlpComplianceRule }
    'pv.labels'            = { Get-Label }
    'pv.labelPolicies'     = { Get-LabelPolicy }
    'pv.autoLabelPolicies' = { Get-AutoSensitivityLabelPolicy }
    'pv.retentionPolicies' = { Get-RetentionCompliancePolicy }
    'pv.retentionRules'    = { Get-RetentionComplianceRule }
}


# ─── Teams ───────────────────────────────────────────────────────────────────
$TeamsAreas = [ordered]@{
    'teams.teams'             = { Get-Team | Select-Object GroupId, Visibility, Archived }
    'teams.federation'        = { Get-CsTenantFederationConfiguration }
    'teams.externalAccess'    = { Get-CsExternalAccessPolicy }
    'teams.client'            = { Get-CsTeamsClientConfiguration }
    'teams.guestMeeting'      = { Get-CsTeamsGuestMeetingConfiguration }
    'teams.guestMessaging'    = { Get-CsTeamsGuestMessagingConfiguration }
    'teams.guestCalling'      = { Get-CsTeamsGuestCallingConfiguration }
    'teams.meetingConfig'     = { Get-CsTeamsMeetingConfiguration }
    'teams.meetingPolicies'   = { Get-CsTeamsMeetingPolicy }
    'teams.messagingPolicies' = { Get-CsTeamsMessagingPolicy }
    'teams.callingPolicies'   = { Get-CsTeamsCallingPolicy }
    'teams.channelPolicies'   = { Get-CsTeamsChannelsPolicy }
    'teams.appSetupPolicies'  = { Get-CsTeamsAppSetupPolicy }
    'teams.appPermPolicies'   = { Get-CsTeamsAppPermissionPolicy }
    'teams.updatePolicies'    = { Get-CsTeamsUpdateManagementPolicy }
}


# Follows @odata.nextLink so collections come back complete.
function Get-GraphAll([string] $Path) {
    $uri   = "$GRAPH$Path"
    $items = [System.Collections.Generic.List[object]]::new()

    while ($uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

        # Single object (e.g. authorizationPolicy), no paging.
        if (-not $page.PSObject.Properties['value']) { return $page }

        $items.AddRange([object[]] $page.value)
        $uri = $page.'@odata.nextLink'
    }

    # Unrolled on purpose: Save-Area wraps the output in @(), empty stays [].
    return $items
}


# Runs one reader and writes its result. Always stores an array, so "nothing
# found" ([]) and "not exported" (no file) stay distinguishable.
function Save-Area([string] $Dir, [string] $Area, [scriptblock] $Read) {
    $file = Join-Path $Dir "$Area.json"

    try {
        $data = @(& $Read)
        ConvertTo-Json -InputObject $data -Depth $JSON_DEPTH -EnumsAsStrings -WarningAction SilentlyContinue |
            Set-Content $file
        Write-Host "  ok   $Area"
        return
    }
    catch {
        $status = $_.Exception.Response.StatusCode.value__
        $reason = $status -eq $HTTP_FORBIDDEN ? 'missing permission' : $_.Exception.Message
        $script:Errors.Add([pscustomobject]@{ area = $Area; status = $status; reason = $reason })

        # A file from an earlier run (e.g. as admin) must not pass as current data.
        if (Test-Path $file) { Remove-Item $file }

        Write-Host "  fail $Area ($reason)" -ForegroundColor Yellow
    }
}


function Save-Areas([string] $Dir, $Areas) {
    foreach ($area in $Areas.Keys) { Save-Area $Dir $area $Areas[$area] }
}


# Checks the license export for an active, user-assigned service plan.
# Unknown (no license export) counts as licensed, so we try rather than skip.
function Test-Plan([string] $Dir, [string] $Plan) {
    $file = Join-Path $Dir 'lic.subscribedSkus.json'
    if (-not (Test-Path $file)) { return $true }

    $skus = Get-Content $file -Raw | ConvertFrom-Json
    $hits = $skus | Where-Object capabilityStatus -eq 'Enabled' |
        ForEach-Object servicePlans |
        Where-Object { $_.servicePlanName -match $Plan -and $_.provisioningStatus -eq 'Success' }

    return [bool] $hits
}


# Tenant GUID via the public OpenID metadata, no sign-in needed. Teams needs the GUID.
# Example issuer: https://login.microsoftonline.com/d5d2f0ab-.../v2.0
function Get-TenantId([string] $Domain) {
    $meta = Invoke-RestMethod "$LOGIN/$Domain/v2.0/.well-known/openid-configuration"
    return ([uri] $meta.issuer).Segments[1].TrimEnd('/')
}


# Without our app, Graph would fall back to the Graph CLI app and its consent.
function Assert-ClientId($Entry) {
    if (-not $Entry.clientId) { throw "clientId missing in tenants.json for '$($Entry.name)'." }
}


# Optional account splat, avoids picking a cached account of another tenant.
function Get-Upn($Entry, [string] $Name) {
    if (-not $Entry.account) { return @{} }

    return @{ $Name = $Entry.account }
}


# Splat for EXO, Purview and Teams: system browser instead of WAM on request.
function Get-BrokerArgs {
    if ($Broker -eq 'Browser') { return @{ DisableWAM = $true } }

    return @{}
}


function Export-Graph($Entry, [string] $Dir) {
    # No -Scopes: the token carries the delegated permissions consented in Entra.
    # Default ContextScope (CurrentUser): with "Process" the WAM token is not
    # reused and the account picker shows twice. No device code: Microsoft
    # recommends blocking it via Conditional Access.
    Connect-MgGraph -TenantId $Entry.tenantId -ClientId $Entry.clientId -NoWelcome

    try {
        # Probe once, so a failed login or missing consent aborts here.
        Invoke-MgGraphRequest -Method GET -Uri "$GRAPH/v1.0/organization" | Out-Null

        foreach ($area in $GraphAreas.Keys) {
            $path = $GraphAreas[$area]
            Save-Area $Dir $area { Get-GraphAll $path }
        }
    }
    finally {
        # Clears the cached context, no token stays behind for the next tenant.
        Disconnect-MgGraph | Out-Null
    }
}


function Export-Exchange($Entry, [string] $Dir) {
    if (-not (Test-Plan $Dir $EXO_PLAN)) { return $NOT_LICENSED }

    $upn    = Get-Upn $Entry 'UserPrincipalName'
    $broker = Get-BrokerArgs
    Connect-ExchangeOnline -Organization $Entry.domain -ShowBanner:$false @upn @broker
    Save-Areas $Dir $ExchangeAreas
    Disconnect-ExchangeOnline -Confirm:$false

    return $EXPORTED
}


function Export-Purview($Entry, [string] $Dir) {
    $upn    = Get-Upn $Entry 'UserPrincipalName'
    $broker = Get-BrokerArgs
    Connect-IPPSSession -Organization $Entry.domain @upn @broker
    Save-Areas $Dir $PurviewAreas
    Disconnect-ExchangeOnline -Confirm:$false

    return $EXPORTED
}


function Export-Teams($Entry, [string] $Dir) {
    if (-not (Test-Plan $Dir $TEAMS_PLAN)) { return $NOT_LICENSED }

    # No -AccountId: in the Teams module it means Windows-integrated sign-in
    # (OS account), not a login hint.
    $broker = Get-BrokerArgs
    Connect-MicrosoftTeams -TenantId $Entry.tenantId @broker | Out-Null
    Save-Areas $Dir $TeamsAreas
    Disconnect-MicrosoftTeams | Out-Null

    return $EXPORTED
}


# Public DNS, no sign-in. Covers custom domains from the Graph export.
# Example row: { domain: schuler.de, mx: mx01.hornetsecurity.com, dmarc: v=DMARC1; p=reject; ... }
function Export-Dns($Entry, [string] $Dir) {
    $file = Join-Path $Dir 'org.domains.json'
    if (-not (Test-Path $file)) { throw 'org.domains.json missing, run Graph first.' }

    $domains = (Get-Content $file -Raw | ConvertFrom-Json) |
        Where-Object { $_.isVerified -and $_.id -notlike '*.onmicrosoft.com' }

    Save-Area $Dir 'dns.records' {
        foreach ($d in $domains) {
            [pscustomobject]@{
                domain = $d.id
                mx     = (Resolve-Dns $d.id 'MX').NameExchange -join ', '
                spf    = (Resolve-Dns $d.id 'TXT').Strings | Where-Object { $_ -like 'v=spf1*' }
                dmarc  = (Resolve-Dns "_dmarc.$($d.id)" 'TXT').Strings -join ''
                dkim1  = (Resolve-Dns "selector1._domainkey.$($d.id)" 'CNAME').NameHost
                dkim2  = (Resolve-Dns "selector2._domainkey.$($d.id)" 'CNAME').NameHost
            }
        }
    }

    return $EXPORTED
}


# Missing records are a finding, not an error.
function Resolve-Dns([string] $Name, [string] $Type) {
    try   { return Resolve-DnsName $Name -Type $Type -DnsOnly | Where-Object Type -eq $Type }
    catch { return $null }
}


function Export-Tenant($Entry) {
    Write-Host "`n== $($Entry.label) ($($Entry.domain))" -ForegroundColor Cyan

    # Before the folder exists: Compare-Inventory takes every folder as an exported tenant.
    try   { Assert-ClientId $Entry }
    catch { Write-Host "  skip ($($_.Exception.Message))" -ForegroundColor Yellow; return }

    $dir = Join-Path $EXPORT_ROOT $Entry.name
    New-Item -ItemType Directory -Force $dir | Out-Null

    $script:Errors = [System.Collections.Generic.List[object]]::new()
    $status = [ordered]@{}

    $Entry | Add-Member -NotePropertyName tenantId -NotePropertyValue (Get-TenantId $Entry.domain) -Force

    foreach ($w in $WORKLOADS) {
        if ($w -notin $Workload) { continue }

        Write-Host "-- $w"

        try {
            $result = & "Export-$w" $Entry $dir
            $status[$w] = $result ?? $EXPORTED
        }
        catch {
            $status[$w] = "failed: $($_.Exception.Message)"
            Write-Host "  fail $w ($($_.Exception.Message))" -ForegroundColor Red
        }
    }

    ConvertTo-Json -InputObject @($script:Errors) | Set-Content (Join-Path $dir '_errors.json')
    ConvertTo-Json -InputObject $status            | Set-Content (Join-Path $dir '_workloads.json')
}


$tenants = Get-Content $TENANT_FILE -Raw | ConvertFrom-Json | Where-Object name -like $Tenant
if (-not $tenants) { throw "No tenant matches '$Tenant'." }

foreach ($t in $tenants) { Export-Tenant $t }
