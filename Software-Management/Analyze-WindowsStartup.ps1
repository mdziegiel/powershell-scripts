<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Analyze-WindowsStartup
SYNOPSIS    : Audit startup apps, services, and scheduled tasks
DESCRIPTION : Audit startup apps, services, and scheduled tasks
================================================================================
#>
[CmdletBinding()]
param([string]$OutputRoot="$env:PUBLIC\Documents\WindowsStartupReports")

Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
$runPath=Join-Path $OutputRoot ("Startup_{0}_{1}" -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd_HHmmss'))
$warnings=New-Object System.Collections.Generic.List[string]

try{
    if($env:OS -ne 'Windows_NT'){throw 'Windows is required.'}
    New-Item $runPath -ItemType Directory -Force|Out-Null

    Get-CimInstance Win32_StartupCommand|
        Select-Object Name,Command,Location,User|
        Export-Csv (Join-Path $runPath 'StartupCommands.csv') -NoTypeInformation

    Get-CimInstance Win32_Service|
        Where-Object StartMode -eq 'Auto'|
        Select-Object Name,DisplayName,State,Status,StartMode,StartName,PathName|
        Export-Csv (Join-Path $runPath 'AutomaticServices.csv') -NoTypeInformation

    try{
        Get-WinEvent -FilterHashtable @{
            LogName='Microsoft-Windows-Diagnostics-Performance/Operational'
            Id=100,101,102,103
            StartTime=(Get-Date).AddDays(-30)
        } -MaxEvents 100 -ErrorAction Stop|
            Select-Object TimeCreated,Id,LevelDisplayName,Message|
            Export-Csv (Join-Path $runPath 'StartupPerformanceEvents.csv') -NoTypeInformation
    }catch{$warnings.Add($_.Exception.Message)}

    $warnings|Out-File (Join-Path $runPath 'Warnings.txt')
    Write-Host "[OK] Startup report created: $runPath" -ForegroundColor Green
    if($warnings.Count -gt 0){exit 2}else{exit 0}
}catch{Write-Error $_.Exception.Message;exit 1}