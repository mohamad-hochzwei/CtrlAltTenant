<#
.SYNOPSIS
    Step 1: builds the comparison matrix from exports/ (rows = settings, columns = tenants).

.DESCRIPTION
        exports/<tenant>/*.json ──► Compare-Inventory.ps1 ──► matrix.xlsx
                                                          ├─ Matrix    settings x tenant, differing rows highlighted
                                                          ├─ Workloads exported / not licensed / failed
                                                          ├─ Licenses  SKU x tenant, consumed/enabled
                                                          ├─ CA        policy name x tenant, state
                                                          ├─ DNS       MX, SPF, DKIM, DMARC per domain
                                                          └─ Errors    failed areas, basis for the permission talk

    n/a = not exported (see Errors) or setting absent. Add a row by adding one entry to $Rows.
#>
$ErrorActionPreference = 'Stop'

$EXPORT_ROOT = Join-Path $PSScriptRoot 'exports'
$TENANT_FILE = Join-Path $PSScriptRoot 'tenants.json'
$OUTPUT      = Join-Path $PSScriptRoot 'matrix.xlsx'
$NO_DATA     = 'n/a'
$DIFF_COLOR  = 'LightYellow'
$GLOBAL      = 'Global'
$MS_TENANTS  = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')   # Microsoft first-party apps
$COPILOT_SKU = 'Copilot'
$REPORT_ONLY = 'enabledForReportingButNotEnforced'

# Entra guest user role ids (fixed by Microsoft).
$GUEST_ROLES = @{
    'a0b1b346-4d3e-4e8b-98f8-753987be4970' = 'same as members'
    '10dae51f-b6af-4016-8d66-8c2a99b929b3' = 'limited (default)'
    '2af84b1e-32c8-42b7-82bc-daa82404023b' = 'restricted'
}

# Microsoft defaults when the setting object does not exist.
$DEFAULT_GROUP_CREATION = 'true'
$DEFAULT_LOCKOUT        = 10


# Loads one export, $null if missing or not readable (see _errors.json).
function Read-Export([string] $Name, [string] $Area) {
    $file = Join-Path $EXPORT_ROOT $Name "$Area.json"
    if (-not (Test-Path $file)) { return $null }

    return Get-Content $file -Raw | ConvertFrom-Json
}


# Reads an area of the tenant currently in the loop ($t, dynamic scope).
function Read-Area([string] $Area) {
    return Read-Export $t.name $Area
}


# Counts items of an area, optionally filtered. $null if the export is missing,
# so the matrix shows n/a instead of a misleading 0.
function Get-AreaCount([string] $Area, [scriptblock] $Filter = { $true }) {
    $file = Join-Path $EXPORT_ROOT $t.name "$Area.json"
    if (-not (Test-Path $file)) { return $null }

    return @(Read-Area $Area | Where-Object $Filter).Count
}


# Tenant-wide default of a Teams/EXO policy set (Identity "Global").
function Get-Global([string] $Area) {
    return Read-Area $Area | Where-Object Identity -eq $GLOBAL
}


# Value of one entry in an Entra groupSettings object, e.g. LockoutThreshold.
function Get-GroupSetting([string] $Template, [string] $Name) {
    $setting = Read-Area 'entra.groupSettings' | Where-Object displayName -eq $Template
    return ($setting.values | Where-Object name -eq $Name).value
}


# "part/total", e.g. "2/3" domains with DMARC.
function Get-Ratio([string] $Area, [scriptblock] $Filter) {
    $all = Get-AreaCount $Area
    if ($null -eq $all) { return $null }

    return "$(Get-AreaCount $Area $Filter)/$all"
}


# Group counts as "A=10, B=2".
function Join-Groups($Items, [string] $Property) {
    if ($null -eq $Items) { return $null }

    return ($Items | Group-Object $Property | Sort-Object Count -Descending | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
}


# Shows "default (x)" when the area was exported but the setting is not set.
function Use-Default([string] $Area, $Value, $Default) {
    if ($null -ne $Value) { return $Value }
    if ($null -eq (Get-AreaCount $Area)) { return $null }

    return "default ($Default)"
}


# Pairs as "A=x, B=y". $null on missing export, avoids a lone "=" that Excel reads as formula.
function Join-Pairs($Items, [string] $Key, [string] $Value) {
    if ($null -eq $Items) { return $null }

    return ($Items | ForEach-Object { "$($_.$Key)=$($_.$Value)" }) -join ', '
}


# Setting ──► extractor. Extractors run once per tenant ($t).
$Rows = [ordered]@{
    # ── Licenses / Copilot
    'Licenses | SKUs active'                   = { Get-AreaCount 'lic.subscribedSkus' { $_.capabilityStatus -eq 'Enabled' } }
    'Copilot | Seats used/available'           = {
        $skus = Read-Area 'lic.subscribedSkus'
        if ($null -eq $skus) { return $null }

        $sku = $skus | Where-Object skuPartNumber -match $COPILOT_SKU
        if ($sku) { "$(($sku.consumedUnits | Measure-Object -Sum).Sum)/$(($sku.prepaidUnits.enabled | Measure-Object -Sum).Sum)" } else { 0 }
    }

    # ── Entra
    'Entra | Members enabled'                  = { Get-AreaCount 'entra.users' { $_.userType -eq 'Member' -and $_.accountEnabled } }
    'Entra | Guests'                           = { Get-AreaCount 'entra.users' { $_.userType -eq 'Guest' } }
    'Entra | Synced from AD'                   = { Get-AreaCount 'entra.users' { $_.onPremisesSyncEnabled } }
    'Entra | Directory sync'                   = {
        $org = Read-Area 'org.organization'
        if ($org) { [bool] $org.onPremisesSyncEnabled }
    }
    'Entra | Groups (M365/total)'              = { Get-Ratio 'entra.groups' { $_.groupTypes -contains 'Unified' } }
    'Entra | Dynamic groups'                   = { Get-AreaCount 'entra.groups' { $_.groupTypes -contains 'DynamicMembership' } }
    'Entra | Group creation by users'          = { Use-Default 'entra.groupSettings' (Get-GroupSetting 'Group.Unified' 'EnableGroupCreation') $DEFAULT_GROUP_CREATION }
    'Entra | Lockout threshold'                = { Use-Default 'entra.groupSettings' (Get-GroupSetting 'Password Rule Settings' 'LockoutThreshold') $DEFAULT_LOCKOUT }
    'Entra | Security defaults'                = { (Read-Area 'entra.securityDefaults').isEnabled }
    'Entra | Users can register apps'          = { (Read-Area 'entra.authorization').defaultUserRolePermissions.allowedToCreateApps }
    'Entra | Users can create tenants'         = { (Read-Area 'entra.authorization').defaultUserRolePermissions.allowedToCreateTenants }
    'Entra | User consent policy'              = {
        $auth = Read-Area 'entra.authorization'
        if (-not $auth) { return $null }

        # Empty list = users cannot consent to apps.
        $ids = $auth.permissionGrantPolicyIdsAssignedToDefaultUserRole | Where-Object { $_ -like 'ManagePermissionGrantsForSelf.*' }
        $ids ? ($ids -replace 'ManagePermissionGrantsForSelf\.', '') -join ', ' : 'disabled'
    }
    'Entra | Admin consent workflow'           = { (Read-Area 'entra.adminConsent').isEnabled }
    'Entra | Guest invite setting'             = { (Read-Area 'entra.authorization').allowInvitesFrom }
    'Entra | Guest user role'                  = {
        $id = (Read-Area 'entra.authorization').guestUserRoleId
        if ($id) { $GUEST_ROLES[$id] ?? $id }
    }
    'Entra | Cross-tenant partners'            = { Get-AreaCount 'entra.crossTenantPartners' }
    'Entra | Device join allowed'              = {
        # "#microsoft.graph.allDeviceRegistrationMembership" ──► "all"
        $type = (Read-Area 'entra.deviceRegistration').azureADJoin.allowedToJoin.'@odata.type'
        if ($type) { $type -replace '#microsoft\.graph\.|DeviceRegistrationMembership', '' }
    }
    'Entra | Device MFA to join'               = { (Read-Area 'entra.deviceRegistration').multiFactorAuthConfiguration }
    'Entra | Auth methods enabled'             = { ((Read-Area 'entra.authMethods').authenticationMethodConfigurations | Where-Object state -eq 'enabled').id -join ', ' }
    'Entra | MFA registered (members)'         = { Get-Ratio 'entra.mfaRegistration' { $_.userType -eq 'member' -and $_.isMfaRegistered } }
    'Entra | Global Admins (active)'           = {
        $gaId = (Read-Area 'entra.roleDefinitions' | Where-Object displayName -eq 'Global Administrator').id
        Get-AreaCount 'entra.roleAssignments' { $_.roleDefinitionId -eq $gaId }
    }
    'Entra | PIM eligible assignments'         = { Get-AreaCount 'entra.pimEligible' }
    'Entra | App registrations'                = { Get-AreaCount 'entra.applications' }
    'Entra | Apps with expired secrets'        = { Get-AreaCount 'entra.applications' { $_.passwordCredentials | Where-Object { [datetime] $_.endDateTime -lt (Get-Date) } } }
    'Entra | Third-party enterprise apps'      = {
        $own = (Read-Area 'org.organization').id
        Get-AreaCount 'entra.servicePrincipals' { $_.servicePrincipalType -eq 'Application' -and $_.appOwnerOrganizationId -notin ($MS_TENANTS + $own) }
    }
    'Entra | Tenant-wide OAuth grants'         = { Get-AreaCount 'entra.oauthGrants' { $_.consentType -eq 'AllPrincipals' } }

    # ── Conditional Access
    'CA | Policies total'                      = { Get-AreaCount 'ca.policies' }
    'CA | Policies enabled'                    = { Get-AreaCount 'ca.policies' { $_.state -eq 'enabled' } }
    'CA | Policies report-only'                = { Get-AreaCount 'ca.policies' { $_.state -eq $REPORT_ONLY } }
    'CA | Named locations'                     = { Get-AreaCount 'ca.namedLocations' }
    'CA | Custom auth strengths'               = { Get-AreaCount 'entra.authStrengths' { $_.policyType -eq 'custom' } }

    # ── Security
    'Security | Secure Score %'                = {
        $s = @(Read-Area 'sec.secureScore')[0]
        if ($s) { [math]::Round(100 * $s.currentScore / $s.maxScore, 1) }
    }

    # ── Intune
    'Intune | Managed devices by OS'           = { Join-Groups (Read-Area 'intune.devices') 'operatingSystem' }
    'Intune | Non-compliant devices'           = { Get-AreaCount 'intune.devices' { $_.complianceState -eq 'noncompliant' } }
    'Intune | Compliance policies'             = { Get-AreaCount 'intune.compliance' }
    'Intune | Compliance unassigned'           = { Get-AreaCount 'intune.compliance' { -not $_.assignments } }
    'Intune | Config profiles (legacy)'        = { Get-AreaCount 'intune.configurations' }
    'Intune | Settings catalog policies'       = { Get-AreaCount 'intune.settingsCatalog' }
    'Intune | Group Policy (ADMX) profiles'    = { Get-AreaCount 'intune.groupPolicy' }
    'Intune | Endpoint security (legacy)'      = { Get-AreaCount 'intune.endpointSecurity' }
    'Intune | Enrollment configurations'       = { Get-AreaCount 'intune.enrollment' }
    'Intune | App protection policies'         = { Get-AreaCount 'intune.appProtection' }
    'Intune | Apps'                            = { Get-AreaCount 'intune.apps' }
    'Intune | Apps assigned'                   = { Get-AreaCount 'intune.apps' { $_.assignments } }
    'Intune | Autopilot profiles'              = { Get-AreaCount 'intune.autopilotProfiles' }
    'Intune | Autopilot devices'               = { Get-AreaCount 'intune.autopilotDevices' }
    'Intune | Feature update profiles'         = { Get-AreaCount 'intune.featureUpdates' }
    'Intune | Platform scripts'                = { Get-AreaCount 'intune.scripts' }
    'Intune | Remediations'                    = { Get-AreaCount 'intune.remediations' }

    # ── Exchange Online
    'Exchange | Mailboxes by type'             = { Join-Pairs (Read-Area 'exo.mailboxCounts') 'Name' 'Count' }
    'Exchange | Modern auth (OAuth2)'          = { (Read-Area 'exo.organization').OAuth2ClientProfileEnabled }
    'Exchange | Mailbox audit disabled'        = { (Read-Area 'exo.organization').AuditDisabled }
    'Exchange | Customer Lockbox'              = { (Read-Area 'exo.organization').CustomerLockBoxEnabled }
    'Exchange | Unified audit log'             = { (Read-Area 'exo.auditConfig').UnifiedAuditLogIngestionEnabled }
    'Exchange | SMTP AUTH disabled'            = { (Read-Area 'exo.transport').SmtpClientAuthenticationDisabled }
    'Exchange | External sender tag'           = { (Read-Area 'exo.externalTagging').Enabled }
    'Exchange | Auto-forward (outbound spam)'  = { (Read-Area 'mdo.outboundSpam' | Where-Object IsDefault).AutoForwardingMode }
    'Exchange | Auto-forward (remote domain)'  = { (Read-Area 'exo.remoteDomains' | Where-Object Identity -eq 'Default').AutoForwardEnabled }
    'Exchange | Accepted domains'              = { Get-AreaCount 'exo.acceptedDomains' }
    'Exchange | Transport rules'               = { Get-AreaCount 'exo.transportRules' }
    'Exchange | Inbound connectors'            = { (Read-Area 'exo.inboundConnectors').Name -join ', ' }
    'Exchange | Outbound connectors'           = { (Read-Area 'exo.outboundConnectors').Name -join ', ' }
    'Exchange | DKIM enabled'                  = { Get-Ratio 'exo.dkim' { $_.Enabled } }
    'Exchange | Authentication policies'       = { Get-AreaCount 'exo.authPolicies' }

    # ── Defender for Office 365
    'Defender | Preset policies (EOP)'         = { Join-Pairs (Read-Area 'mdo.presetEop') 'Identity' 'State' }
    'Defender | Preset policies (MDO)'         = { Join-Pairs (Read-Area 'mdo.presetAtp') 'Identity' 'State' }
    'Defender | Anti-phish policies'           = { Get-AreaCount 'mdo.antiPhish' }
    'Defender | Anti-spam policies'            = { Get-AreaCount 'mdo.antiSpam' }
    'Defender | Anti-malware policies'         = { Get-AreaCount 'mdo.antiMalware' }
    'Defender | Safe Links policies'           = { Get-AreaCount 'mdo.safeLinks' }
    'Defender | Safe Attachments policies'     = { Get-AreaCount 'mdo.safeAttachments' }
    'Defender | MDO for SPO/OD/Teams'          = { (Read-Area 'mdo.atpGlobal').EnableATPForSPOTeamsODB }
    'Defender | Safe Documents'                = { (Read-Area 'mdo.atpGlobal').EnableSafeDocs }

    # ── Purview
    'Purview | DLP policies (enabled/total)'   = { Get-Ratio 'pv.dlpPolicies' { $_.Mode -eq 'Enable' } }
    'Purview | Sensitivity labels'             = { Get-AreaCount 'pv.labels' }
    'Purview | Label policies'                 = { Get-AreaCount 'pv.labelPolicies' }
    'Purview | Auto-label policies'            = { Get-AreaCount 'pv.autoLabelPolicies' }
    'Purview | Retention policies'             = { Get-AreaCount 'pv.retentionPolicies' }

    # ── Teams
    'Teams | Teams'                            = { Get-AreaCount 'teams.teams' }
    'Teams | Teams archived'                   = { Get-AreaCount 'teams.teams' { $_.Archived } }
    'Teams | External access (federation)'     = { (Read-Area 'teams.federation').AllowFederatedUsers }
    'Teams | Allowed domains'                  = { "$((Read-Area 'teams.federation').AllowedDomains)" }
    'Teams | Teams personal accounts'          = { (Read-Area 'teams.federation').AllowTeamsConsumer }
    'Teams | Guest access'                     = { (Get-Global 'teams.client').AllowGuestUser }
    'Teams | Third-party storage'              = {
        $c = Get-Global 'teams.client'
        if ($c) { ('Dropbox', 'Box', 'GoogleDrive', 'ShareFile', 'Egnyte' | Where-Object { $c."Allow$_" }) -join ', ' }
    }
    'Teams | Anonymous join (org)'             = {
        $cfg = Read-Area 'teams.meetingConfig'
        if ($cfg) { -not $cfg.DisableAnonymousJoin }
    }
    'Teams | Anonymous join (policy)'          = { (Get-Global 'teams.meetingPolicies').AllowAnonymousUsersToJoinMeeting }
    'Teams | Lobby bypass'                     = { (Get-Global 'teams.meetingPolicies').AutoAdmittedUsers }
    'Teams | Cloud recording'                  = { (Get-Global 'teams.meetingPolicies').AllowCloudRecording }
    'Teams | Transcription'                    = { (Get-Global 'teams.meetingPolicies').AllowTranscription }
    'Teams | External give control'            = { (Get-Global 'teams.meetingPolicies').AllowExternalParticipantGiveRequestControl }
    'Teams | Meeting policies'                 = { Get-AreaCount 'teams.meetingPolicies' }
    'Teams | Messaging policies'               = { Get-AreaCount 'teams.messagingPolicies' }
    'Teams | App setup policies'               = { Get-AreaCount 'teams.appSetupPolicies' }
    'Teams | Private channels'                 = { (Get-Global 'teams.channelPolicies').AllowPrivateChannelCreation }

    # ── SharePoint / OneDrive
    'SharePoint | External sharing'            = { (Read-Area 'spo.settings').sharingCapability }
    'SharePoint | Domain restriction'          = { (Read-Area 'spo.settings').sharingDomainRestrictionMode }
    'SharePoint | Guests can reshare'          = { (Read-Area 'spo.settings').isResharingByExternalUsersEnabled }
    'SharePoint | Sync from unmanaged'         = { (Read-Area 'spo.settings').isUnmanagedSyncAppForTenantRestricted }
    'SharePoint | Idle session sign-out'       = { (Read-Area 'spo.settings').idleSessionSignOut.isEnabled }
    'SharePoint | Loop enabled'                = { (Read-Area 'spo.settings').isLoopEnabled }
    'OneDrive | Default storage (MB)'          = { (Read-Area 'spo.settings').personalSiteDefaultStorageLimitInMB }

    # ── DNS
    'DNS | Custom domains'                     = { Get-AreaCount 'dns.records' }
    'DNS | MX hosts'                           = { ((Read-Area 'dns.records').mx | Sort-Object -Unique) -join ', ' }
    'DNS | SPF present'                        = { Get-Ratio 'dns.records' { $_.spf } }
    'DNS | DKIM CNAMEs present'                = { Get-Ratio 'dns.records' { $_.dkim1 -and $_.dkim2 } }
    'DNS | DMARC enforcing'                    = { Get-Ratio 'dns.records' { $_.dmarc -match 'p=(reject|quarantine)' } }
}


$tenants = Get-Content $TENANT_FILE -Raw | ConvertFrom-Json | Where-Object { Test-Path (Join-Path $EXPORT_ROOT $_.name) }
if (-not $tenants) { throw 'No exports found. Run Export-Inventory.ps1 first.' }


# Matrix: one row per setting, one column per tenant.
$matrix = foreach ($setting in $Rows.Keys) {
    $area, $name = $setting -split ' \| '
    $row = [ordered]@{ Area = $area; Setting = $name }

    foreach ($t in $tenants) {
        $value = try { & $Rows[$setting] } catch { $null }
        $isNumber = $value -is [ValueType] -and $value -isnot [bool]

        $row[$t.name] = $null -eq $value -or "$value" -eq '' ? $NO_DATA : ($isNumber ? $value : "$value")
    }

    [pscustomobject] $row
}


# Workloads: did each part of the export run, and was the workload licensed.
$workloads = foreach ($w in 'Graph', 'Exchange', 'Purview', 'Teams', 'Dns') {
    $row = [ordered]@{ Workload = $w }
    foreach ($t in $tenants) { $row[$t.name] = (Read-Export $t.name '_workloads').$w ?? $NO_DATA }
    [pscustomobject] $row
}


# Licenses: consumed/enabled per SKU and tenant.
$skus = foreach ($t in $tenants) {
    foreach ($s in @(Read-Export $t.name 'lic.subscribedSkus')) {
        if (-not $s) { continue }
        [pscustomobject]@{ Tenant = $t.name; Sku = $s.skuPartNumber; Seats = "$($s.consumedUnits)/$($s.prepaidUnits.enabled)" }
    }
}

$licenses = foreach ($sku in ($skus.Sku | Sort-Object -Unique)) {
    $row = [ordered]@{ Sku = $sku }
    foreach ($t in $tenants) { $row[$t.name] = ($skus | Where-Object { $_.Tenant -eq $t.name -and $_.Sku -eq $sku }).Seats }
    [pscustomobject] $row
}


# CA: policy names side by side, makes overlaps and gaps visible.
$policies = foreach ($t in $tenants) {
    foreach ($p in @(Read-Export $t.name 'ca.policies')) {
        if (-not $p) { continue }
        [pscustomobject]@{ Tenant = $t.name; Policy = $p.displayName; State = $p.state }
    }
}

$ca = foreach ($policy in ($policies.Policy | Sort-Object -Unique)) {
    $row = [ordered]@{ Policy = $policy }
    foreach ($t in $tenants) { $row[$t.name] = ($policies | Where-Object { $_.Tenant -eq $t.name -and $_.Policy -eq $policy }).State }
    [pscustomobject] $row
}


# DNS and Errors: long format, one row per domain / failed area.
$dns = foreach ($t in $tenants) {
    @(Read-Export $t.name 'dns.records') | Where-Object { $_ } | Select-Object @{ n = 'Tenant'; e = { $t.name } }, *
}

$errors = foreach ($t in $tenants) {
    @(Read-Export $t.name '_errors') | Where-Object { $_ } | Select-Object @{ n = 'Tenant'; e = { $t.name } }, *
}


Remove-Item $OUTPUT -ErrorAction SilentlyContinue

$xl = $matrix | Export-Excel $OUTPUT -WorksheetName 'Matrix' -AutoSize -FreezeTopRow -BoldTopRow -NoNumberConversion * -PassThru
$sheet = $xl.Workbook.Worksheets['Matrix']

# Highlight rows whose tenant values differ.
for ($i = 0; $i -lt $matrix.Count; $i++) {
    $values = $tenants | ForEach-Object { $matrix[$i].($_.name) } | Sort-Object -Unique
    if (@($values).Count -le 1) { continue }

    $excelRow = $i + 2
    $sheet.Row($excelRow).Style.Fill.PatternType = 'Solid'
    $sheet.Row($excelRow).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::$DIFF_COLOR)
}

Close-ExcelPackage $xl

$sheets = [ordered]@{ Workloads = $workloads; Licenses = $licenses; CA = $ca; DNS = $dns; Errors = $errors }

foreach ($name in $sheets.Keys) {
    if (-not $sheets[$name]) { continue }
    $sheets[$name] | Export-Excel $OUTPUT -WorksheetName $name -AutoSize -FreezeTopRow -BoldTopRow -NoNumberConversion *
}

Write-Host "Written: $OUTPUT"
