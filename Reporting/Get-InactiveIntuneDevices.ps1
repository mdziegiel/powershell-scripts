#Requires -Version 5.1
<#
.SYNOPSIS
    Reports all Intune-managed devices not seen for X days.
.DESCRIPTION
    Connects to Microsoft Graph and exports devices that haven't checked in
    within the specified number of days. Useful for identifying stale or
    abandoned endpoints that may need cleanup or follow-up.
.PARAMETER DaysInactive
    Number of days since last sync to flag a device. Default is 30.
.PARAMETER OutputPath
    Path for the CSV output file. Defaults to script directory.
.EXAMPLE
    .\Get-InactiveIntuneDevices.ps1
    .\Get-InactiveIntuneDevices.ps1 -DaysInactive 60 -OutputPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [int]$DaysInactive = 30,
    [string]$OutputPath = $PSScriptRoot
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "InactiveIntuneDevices_$Timestamp.csv"
$InactiveThreshold = (Get-Date).AddDays(-$DaysInactive).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Connect to Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All" -NoWelcome -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

Write-Host "Querying Intune for devices inactive since $DaysInactive days ago..." -ForegroundColor Cyan

$Devices = Get-MgDeviceManagementManagedDevice -All -Filter "lastSyncDateTime le $InactiveThreshold" -Property `
    DeviceName, UserDisplayName, UserPrincipalName, OperatingSystem, OsVersion,
    ComplianceState, ManagementState, LastSyncDateTime, EnrolledDateTime,
    SerialNumber, Model, Manufacturer, ManagedDeviceOwnerType |
    Select-Object DeviceName, UserDisplayName, UserPrincipalName,
        OperatingSystem, OsVersion, ComplianceState, ManagementState,
        SerialNumber, Model, Manufacturer, ManagedDeviceOwnerType,
        LastSyncDateTime, EnrolledDateTime,
        @{N="DaysSinceSync"; E={
            if ($_.LastSyncDateTime) { (New-TimeSpan -Start $_.LastSyncDateTime -End (Get-Date)).Days }
            else { "Never" }
        }}

$Devices | Sort-Object LastSyncDateTime |
    Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nFound $($Devices.Count) inactive devices (not seen in $DaysInactive+ days)." -ForegroundColor Yellow
Write-Host "Report saved to: $ReportPath" -ForegroundColor Green

Disconnect-MgGraph | Out-Null
