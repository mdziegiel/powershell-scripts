#Requires -Version 5.1
<#
.SYNOPSIS
    Audits disabled user accounts in Active Directory (hk.lan).
.DESCRIPTION
    Exports all disabled user accounts with last logon, OU location, and group memberships.
    Useful for identifying stale accounts that should be cleaned up or reviewed.
.PARAMETER OutputPath
    Path for the CSV output file. Defaults to script directory.
.PARAMETER DaysSinceLogon
    Filter to only show accounts not logged in for X days. Default is 0 (all disabled accounts).
.EXAMPLE
    .\Get-DisabledAccountAudit.ps1
    .\Get-DisabledAccountAudit.ps1 -OutputPath "C:\Reports" -DaysSinceLogon 90
#>

[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [int]$DaysSinceLogon = 0
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "DisabledAccountAudit_$Timestamp.csv"

Write-Host "Querying disabled accounts in hk.lan..." -ForegroundColor Cyan

$Users = Get-ADUser -Filter { Enabled -eq $false } -Properties `
    DisplayName, SamAccountName, UserPrincipalName, DistinguishedName,
    LastLogonDate, WhenCreated, WhenChanged, MemberOf, Description, Title, Department |
    Select-Object DisplayName, SamAccountName, UserPrincipalName,
        @{N="OU"; E={ ($_.DistinguishedName -split ",",2)[1] }},
        LastLogonDate, WhenCreated, WhenChanged, Description, Title, Department,
        @{N="DaysSinceLogon"; E={
            if ($_.LastLogonDate) { (New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).Days }
            else { "Never" }
        }},
        @{N="GroupMemberships"; E={ ($_.MemberOf | ForEach-Object { ($_ -split ",")[0] -replace "CN=" }) -join "; " }}

# Apply days filter if specified
if ($DaysSinceLogon -gt 0) {
    $Users = $Users | Where-Object { $_.DaysSinceLogon -ne "Never" -and [int]$_.DaysSinceLogon -ge $DaysSinceLogon }
}

$Users | Sort-Object LastLogonDate |
    Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nFound $($Users.Count) disabled accounts." -ForegroundColor Green
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow
