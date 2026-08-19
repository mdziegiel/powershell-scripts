<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Setup-M365EOC-AppOnly
VERSION     : 3.0
PURPOSE     : One-time setup for Microsoft 365 Enterprise Operations Center v3.0.
              Creates a certificate, creates an Entra app registration, grants
              Microsoft Graph application permissions, and writes the dashboard
              config file.

RUN AS      : Normal PowerShell 7 window. Sign in with a Global Administrator or
              an account allowed to create app registrations and grant consent.
==============================================================================
#>

$ErrorActionPreference = 'Stop'
$ReportPath = 'C:\Reporting'
$ConfigPath = Join-Path $ReportPath 'M365EOC_AppOnly_Config.json'
$CertExportPath = Join-Path $ReportPath 'M365EOC_AppOnly_PublicCert.cer'
$AppDisplayName = 'M365 Enterprise Operations Center'

New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    $m = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $m) {
        Write-Host "Installing $Name..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -Force
}

Ensure-Module Microsoft.Graph.Authentication
Ensure-Module Microsoft.Graph.Applications

Write-Host "Connecting to Microsoft Graph for one-time app registration setup..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','Directory.Read.All' -NoWelcome
$ctx = Get-MgContext
$TenantId = $ctx.TenantId

Write-Host "Creating certificate in CurrentUser\My..." -ForegroundColor Cyan
$cert = New-SelfSignedCertificate -Subject 'CN=M365EnterpriseOperationsCenter' -CertStoreLocation 'Cert:\CurrentUser\My' -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
Export-Certificate -Cert $cert -FilePath $CertExportPath -Force | Out-Null

$keyCredential = @{
    Type = 'AsymmetricX509Cert'
    Usage = 'Verify'
    Key = $cert.RawData
    DisplayName = 'M365EOC App-Only Certificate'
    StartDateTime = $cert.NotBefore
    EndDateTime = $cert.NotAfter
}

Write-Host "Creating Entra app registration: $AppDisplayName" -ForegroundColor Cyan
$app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience 'AzureADMyOrg' -KeyCredentials @($keyCredential)
$sp = New-MgServicePrincipal -AppId $app.AppId
Start-Sleep -Seconds 3

$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ConsistencyLevel eventual

$requiredGraphAppPermissions = @(
    'User.Read.All',
    'Directory.Read.All',
    'Application.Read.All',
    'RoleManagement.Read.Directory',
    'Policy.Read.All',
    'Organization.Read.All',
    'LicenseAssignment.Read.All',
    'ServiceHealth.Read.All',
    'ServiceMessage.Read.All',
    'AuditLog.Read.All',
    'SecurityEvents.Read.All',
    'IdentityRiskyUser.Read.All',
    'Reports.Read.All',
    'UserAuthenticationMethod.Read.All',
    'DeviceManagementManagedDevices.Read.All',
    'DeviceManagementConfiguration.Read.All'
)

Write-Host "Granting Microsoft Graph application permissions..." -ForegroundColor Cyan
foreach ($perm in $requiredGraphAppPermissions) {
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
    if ($null -eq $role) {
        Write-Warning "Could not find Graph application permission: $perm"
        continue
    }
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $role.Id | Out-Null
        Write-Host "Granted: $perm" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -match 'Permission being assigned already exists') {
            Write-Host "Already granted: $perm" -ForegroundColor DarkGreen
        }
        else {
            Write-Warning "Failed to grant $perm. $($_.Exception.Message)"
        }
    }
}

$config = [ordered]@{
    TenantId = $TenantId
    ClientId = $app.AppId
    CertificateThumbprint = $cert.Thumbprint
    CertificateStore = 'CurrentUser\My'
    CreatedBy = $ctx.Account
    CreatedOn = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}
$config | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green
Write-Host "Config file: $ConfigPath" -ForegroundColor Green
Write-Host "Public certificate exported to: $CertExportPath" -ForegroundColor Green
Write-Host "ClientId: $($app.AppId)" -ForegroundColor Green
Write-Host "TenantId: $TenantId" -ForegroundColor Green
Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
Write-Host ""
Write-Host "Next: run M365EnterpriseOperationsCenter_v3_0.ps1" -ForegroundColor Cyan
