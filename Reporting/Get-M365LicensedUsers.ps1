<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get M365 Licensed Users
SYNOPSIS    : Reports Microsoft 365 licensed users and assigned licenses
DESCRIPTION : Connects to Microsoft Graph and exports users with assigned
              Microsoft 365 licenses, including license details, account
              status, and sign-in information. Useful for license auditing
              and cost analysis
==============================================================================
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
