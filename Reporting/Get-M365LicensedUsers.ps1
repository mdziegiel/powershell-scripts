#Requires -Version 5.1
<#
.SYNOPSIS
    Exports all Microsoft 365 licensed users and their assigned licenses.
.DESCRIPTION
    Connects to Microsoft Graph and exports all users with M365 licenses,
    including license SKU names, account status, and last sign-in time.
    Useful for license audits and cost reviews.
.PARAMETER OutputPath
    Path for the CSV output file. Defaults to script directory.
.PARAMETER LicensedOnly
    Switch to export only users who have at least one license assigned. Default: true.
.EXAMPLE
    .\Get-M365LicensedUsers.ps1
    .\Get-M365LicensedUsers.ps1 -OutputPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "M365LicensedUsers_$Timestamp.csv"

# Connect to Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# Get SKU friendly names
Write-Host "Retrieving license SKU definitions..." -ForegroundColor Cyan
$SKUs = Get-MgSubscribedSku -All | Select-Object SkuId, SkuPartNumber,
    @{N="FriendlyName"; E={
        $SkuMap = @{
            "SPE_E3"           = "Microsoft 365 E3"
            "SPE_E5"           = "Microsoft 365 E5"
            "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
            "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
            "SPB"              = "Microsoft 365 Business Premium"
            "EXCHANGESTANDARD" = "Exchange Online Plan 1"
            "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
            "POWER_BI_PRO"     = "Power BI Pro"
            "PROJECTPREMIUM"   = "Project Plan 5"
            "VISIOCLIENT"      = "Visio Plan 2"
        }
        if ($SkuMap[$_.SkuPartNumber]) { $SkuMap[$_.SkuPartNumber] } else { $_.SkuPartNumber }
    }}

$SKULookup = @{}
foreach ($SKU in $SKUs) { $SKULookup[$SKU.SkuId] = $SKU.FriendlyName }

# Get licensed users
Write-Host "Querying licensed users..." -ForegroundColor Cyan
$Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable ignored `
    -Property DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses,
              Department, JobTitle, CreatedDateTime, SignInActivity |
    Select-Object DisplayName, UserPrincipalName, AccountEnabled, Department, JobTitle, CreatedDateTime,
        @{N="LastSignIn"; E={ $_.SignInActivity.LastSignInDateTime }},
        @{N="DaysSinceSignIn"; E={
            if ($_.SignInActivity.LastSignInDateTime) {
                (New-TimeSpan -Start $_.SignInActivity.LastSignInDateTime -End (Get-Date)).Days
            } else { "Never" }
        }},
        @{N="AssignedLicenses"; E={
            ($_.AssignedLicenses | ForEach-Object {
                if ($SKULookup[$_.SkuId]) { $SKULookup[$_.SkuId] } else { $_.SkuId }
            }) -join "; "
        }},
        @{N="LicenseCount"; E={ $_.AssignedLicenses.Count }}

$Users | Sort-Object DisplayName |
    Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nExported $($Users.Count) licensed users." -ForegroundColor Green
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow

Disconnect-MgGraph | Out-Null
