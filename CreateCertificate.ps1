<#
.SYNOPSIS
    One certificate per tenant for the app-only inventory.

.DESCRIPTION
    The private key is created non-exportable in the current user's store: it
    never exists as a file and cannot be copied to another machine. Only the
    public key is written, as input for CreateEntraApp.ps1.

        CreateCertificate.ps1 -Tenant SyTech
          ├─ Cert:\CurrentUser\My   CN=hochzwei-Inventory-SyTech, private key, non-exportable
          └─ certs/SyTech.cer       public key ──► CreateEntraApp.ps1 -CertificatePath

    A breached key affects one tenant only. Delete the certificate when the
    project ends or the tenant is done.

.EXAMPLE
    ./CreateCertificate.ps1 -Tenant SyTech
    ./CreateCertificate.ps1 -Tenant SyTech -Months 3
#>
param(
    # Tenant name as in tenants.json.
    [Parameter(Mandatory)]
    [string] $Tenant,

    # Short lifetime limits the damage of a leaked key.
    [ValidateRange(1, 12)]
    [int] $Months = 6
)

$ErrorActionPreference = 'Stop'

$CERT_STORE  = 'Cert:\CurrentUser\My'
$CERT_DIR    = Join-Path $PSScriptRoot 'certs'
$TENANT_FILE = Join-Path $PSScriptRoot 'tenants.json'
$SUBJECT     = "CN=hochzwei-Inventory-$Tenant"

$entry = Get-Content $TENANT_FILE -Raw | ConvertFrom-Json | Where-Object name -eq $Tenant
if (-not $entry) { throw "No tenant '$Tenant' in tenants.json." }

# A second certificate with the same subject makes it unclear which one is in Entra.
$existing = Get-ChildItem $CERT_STORE | Where-Object Subject -eq $SUBJECT
if ($existing) {
    throw "Certificate $SUBJECT exists already ($($existing.Thumbprint -join ', ')). Remove it first: Remove-Item $CERT_STORE\<thumbprint>"
}

$cert = New-SelfSignedCertificate `
    -Subject           $SUBJECT `
    -CertStoreLocation $CERT_STORE `
    -KeyExportPolicy   NonExportable `
    -KeySpec           Signature `
    -KeyAlgorithm      RSA `
    -KeyLength         2048 `
    -HashAlgorithm     SHA256 `
    -NotAfter          (Get-Date).AddMonths($Months)

New-Item -ItemType Directory -Force $CERT_DIR | Out-Null
$cerPath = Join-Path $CERT_DIR "$Tenant.cer"
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

[pscustomobject]@{
    Tenant      = $Tenant
    Subject     = $cert.Subject
    certificate = $cert.Thumbprint
    expires     = $cert.NotAfter.ToString('yyyy-MM-dd')
    publicKey   = $cerPath
} | Format-List

Write-Host "Next: ./CreateEntraApp.ps1 -TenantId $($entry.domain) -AppName 'hochzwei Inventory' -CertificatePath $cerPath"
