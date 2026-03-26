<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Set DNS Suffix Search List
SYNOPSIS    : Configures the DNS suffix search list on the local machine
DESCRIPTION : Sets the DNS suffix search list in the registry and on all
              active network adapters. Useful for Entra-joined or workgroup
              devices that need to resolve internal resources without
              requiring fully qualified domain names
==============================================================================
#>

#Requires -Version 5.1
<#
.SYNOPSIS
    Configures the DNS suffix search list on the local machine.
.DESCRIPTION
    Sets the DNS suffix search list in the registry and on all active network adapters.
    Useful for Entra-joined or workgroup machines that need to resolve hk.lan resources
    without typing the full FQDN.
.PARAMETER Suffixes
    Array of DNS suffixes to set. Defaults to hk.lan suffixes.
.PARAMETER Append
    Switch to append suffixes to the existing list instead of replacing it.
.EXAMPLE
    .\Set-DNSSuffixSearchList.ps1
    .\Set-DNSSuffixSearchList.ps1 -Suffixes "hk.lan","corp.hk.lan"
    .\Set-DNSSuffixSearchList.ps1 -Suffixes "hk.lan" -Append
#>

[CmdletBinding()]
param(
    [string[]]$Suffixes = @("hk.lan", "corp.hk.lan"),
    [switch]$Append
)

# Requires elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

$RegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"

# Get existing suffixes if appending
if ($Append) {
    $Existing = (Get-ItemProperty -Path $RegPath -Name "SearchList" -ErrorAction SilentlyContinue).SearchList
    if ($Existing) {
        $ExistingList = $Existing -split ","
        $Suffixes = ($ExistingList + $Suffixes | Select-Object -Unique)
    }
}

$SuffixString = $Suffixes -join ","

Write-Host "Setting DNS suffix search list to: $SuffixString" -ForegroundColor Cyan

# Set in registry (applies globally)
Set-ItemProperty -Path $RegPath -Name "SearchList" -Value $SuffixString -Type String -Force
Write-Host "Registry updated." -ForegroundColor Green

# Apply to all active adapters
$Adapters = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
foreach ($Adapter in $Adapters) {
    $Result = $Adapter.SetDNSDomain($Suffixes[0])
    $Result2 = $Adapter.SetDNSSuffixSearchOrder($Suffixes)
    Write-Host "Adapter '$($Adapter.Description)': Domain=$($Result.ReturnValue) SearchOrder=$($Result2.ReturnValue)" -ForegroundColor $(if ($Result.ReturnValue -eq 0) { "Green" } else { "Yellow" })
}

Write-Host "`nDNS suffix search list configured. A reboot may be required for full effect." -ForegroundColor Yellow
