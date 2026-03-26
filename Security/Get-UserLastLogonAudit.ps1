<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get User Last Logon Audit
SYNOPSIS    : Reports last logon times for Active Directory users
DESCRIPTION : Queries Active Directory for user accounts and exports
              last logon date, account status, and key attributes.
              Supports filtering for inactive users and enabled accounts
              to assist with auditing and cleanup
==============================================================================
#>

[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [int]$InactiveDays = 90,
    [switch]$EnabledOnly
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "UserLastLogonAudit_$Timestamp.csv"
$InactiveThreshold = (Get-Date).AddDays(-$InactiveDays)

Write-Host "Querying all user accounts in hk.lan..." -ForegroundColor Cyan

$Filter = if ($EnabledOnly) { { Enabled -eq $true } } else { "*" }

$Users = Get-ADUser -Filter $Filter -Properties `
    DisplayName, SamAccountName, UserPrincipalName, Enabled,
    LastLogonDate, WhenCreated, PasswordLastSet, PasswordNeverExpires,
    Title, Department, Manager, DistinguishedName |
    Select-Object DisplayName, SamAccountName, UserPrincipalName,
        @{N="Enabled"; E={ $_.Enabled }},
        Title, Department,
        @{N="Manager"; E={ if ($_.Manager) { (Get-ADUser $_.Manager).DisplayName } else { "" } }},
        @{N="OU"; E={ ($_.DistinguishedName -split ",",2)[1] }},
        LastLogonDate, WhenCreated, PasswordLastSet, PasswordNeverExpires,
        @{N="DaysSinceLogon"; E={
            if ($_.LastLogonDate) { (New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).Days }
            else { "Never Logged In" }
        }},
        @{N="InactiveFlag"; E={
            if ($_.LastLogonDate -and $_.LastLogonDate -lt $InactiveThreshold) { "INACTIVE" }
            elseif (-not $_.LastLogonDate) { "NEVER LOGGED IN" }
            else { "" }
        }}

$Users | Sort-Object LastLogonDate |
    Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

$InactiveCount = ($Users | Where-Object { $_.InactiveFlag -ne "" }).Count
Write-Host "`nTotal accounts exported: $($Users.Count)" -ForegroundColor Green
Write-Host "Accounts inactive/never logged in: $InactiveCount" -ForegroundColor Yellow
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow
