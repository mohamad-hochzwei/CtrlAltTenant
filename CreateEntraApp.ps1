param(
    # Tenant domain or GUID.
    [Parameter(Mandatory)]
    [string] $TenantId,

    [Parameter(Mandatory)]
    [string] $AppName,

    # Public key only (.cer). Never pass the .pfx.
    [Parameter(Mandatory)]
    [string] $CertificatePath,

    # Graph application permissions, read-only. Same set as the export needs.
    [string[]] $Permissions = @(
        'Organization.Read.All'
        'Directory.Read.All'
        'Policy.Read.All'
        'Policy.Read.DeviceConfiguration'
        'RoleManagement.Read.Directory'
        'AuditLog.Read.All'
        'SecurityEvents.Read.All'
        'DeviceManagementConfiguration.Read.All'
        'DeviceManagementApps.Read.All'
        'DeviceManagementManagedDevices.Read.All'
        'DeviceManagementServiceConfig.Read.All'
        'DeviceManagementScripts.Read.All'
        'SharePointTenantSettings.Read.All'
    )
)

$ErrorActionPreference = 'Stop'

$GRAPH            = 'https://graph.microsoft.com/v1.0'
$GRAPH_APP_ID     = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph, same in every tenant
$EXCHANGE_APP_ID  = '00000002-0000-0ff1-ce00-000000000000'   # Office 365 Exchange Online
$EXCHANGE_PERMS   = @('Exchange.ManageAsApp')                # EXO, Defender, Purview
$GLOBAL_READER    = 'f2ef992c-3afb-46b9-b7cf-a126ee74c451'   # role template id, fixed by Microsoft
$RETRY_COUNT      = 6
$RETRY_SECONDS    = 10

# Write rights for the setup only. The app itself gets read-only permissions.
$SetupScopes = @(
    'Application.ReadWrite.All'           # create the app registration
    'AppRoleAssignment.ReadWrite.All'     # grant application permissions
    'RoleManagement.ReadWrite.Directory'  # assign Global Reader
)


function Invoke-Graph([string] $Method, [string] $Path, $Body) {
    $request = @{ Method = $Method; Uri = "$GRAPH$Path"; OutputType = 'PSObject' }
    if ($Body) { $request.Body = ConvertTo-Json $Body -Depth 10; $request.ContentType = 'application/json' }

    return Invoke-MgGraphRequest @request
}


# New service principals take a moment to replicate; early calls fail with "does not exist".
function Invoke-WithRetry([scriptblock] $Action) {
    for ($i = 1; ; $i++) {
        try { return & $Action }
        catch {
            if ($i -ge $RETRY_COUNT) { throw }
            Write-Host "  waiting for replication ($i/$RETRY_COUNT)" -ForegroundColor DarkGray
            Start-Sleep -Seconds $RETRY_SECONDS
        }
    }
}


# Permission names: app role ids of the resource (Graph, Exchange). Throws on
# names that do not exist as application permission.
function Resolve-AppRoles([string] $ResourceAppId, [string[]] $Names) {
    $sp = Invoke-Graph GET "/servicePrincipals(appId='$ResourceAppId')?`$select=id,displayName,appRoles"

    $roles = foreach ($name in $Names) {
        $role = $sp.appRoles | Where-Object { $_.value -eq $name -and $_.allowedMemberTypes -contains 'Application' -and $_.isEnabled }
        if (-not $role) { throw "'$name' is not an application permission of $($sp.displayName)." }
        [pscustomobject]@{ Name = $name; Id = $role.id }
    }

    return [pscustomobject]@{ AppId = $ResourceAppId; SpId = $sp.id; Roles = @($roles) }
}


$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path $CertificatePath).Path)
if ($cert.HasPrivateKey) { throw 'Certificate contains a private key. Pass the .cer, not the .pfx.' }

Connect-MgGraph -TenantId $TenantId -Scopes $SetupScopes -NoWelcome

try {
    # ── Check first, create nothing on error
    $resources = @(
        Resolve-AppRoles $GRAPH_APP_ID $Permissions
        Resolve-AppRoles $EXCHANGE_APP_ID $EXCHANGE_PERMS
    )

    $existing = Invoke-Graph GET "/applications?`$filter=displayName eq '$($AppName -replace "'", "''")'&`$select=id"
    if ($existing.value) { throw "App '$AppName' already exists in this tenant. Delete it or choose another name." }

    # ── App registration with certificate and requested permissions
    $app = Invoke-Graph POST '/applications' @{
        displayName            = $AppName
        signInAudience         = 'AzureADMyOrg'
        requiredResourceAccess = @($resources | ForEach-Object {
            @{ resourceAppId = $_.AppId; resourceAccess = @($_.Roles | ForEach-Object { @{ id = $_.Id; type = 'Role' } }) }
        })
        keyCredentials         = @(@{
            type          = 'AsymmetricX509Cert'
            usage         = 'Verify'
            key           = [Convert]::ToBase64String($cert.RawData)
            displayName   = $cert.Subject
            startDateTime = $cert.NotBefore.ToUniversalTime().ToString('o')
            endDateTime   = $cert.NotAfter.ToUniversalTime().ToString('o')
        })
    }
    Write-Host "  ok   app registration $($app.appId)"

    $sp = Invoke-WithRetry { Invoke-Graph POST '/servicePrincipals' @{ appId = $app.appId } }
    Write-Host "  ok   service principal $($sp.id)"

    # Admin consent: one app role assignment per permission
    foreach ($resource in $resources) {
        foreach ($role in $resource.Roles) {
            $grant = @{ principalId = $sp.id; resourceId = $resource.SpId; appRoleId = $role.Id }
            Invoke-WithRetry { Invoke-Graph POST "/servicePrincipals/$($sp.id)/appRoleAssignments" $grant } | Out-Null
            Write-Host "  ok   $($role.Name)"
        }
    }

    # Global Reader: Exchange, Purview and Teams check the directory role
    $assignment = @{ principalId = $sp.id; roleDefinitionId = $GLOBAL_READER; directoryScopeId = '/' }
    Invoke-WithRetry { Invoke-Graph POST '/roleManagement/directory/roleAssignments' $assignment } | Out-Null
    Write-Host '  ok   Global Reader'

    [pscustomobject]@{
        TenantId    = (Get-MgContext).TenantId
        AppName     = $AppName
        clientId    = $app.appId
        certificate = $cert.Thumbprint
        expires     = $cert.NotAfter.ToString('yyyy-MM-dd')
    } | Format-List
}
finally {
    Disconnect-MgGraph | Out-Null
}
