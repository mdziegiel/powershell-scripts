<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : check-health
SYNOPSIS    : This PowerShell script queries the system health of the local
              computer (hardware, software, and network) and
              prints it.
DESCRIPTION : This PowerShell script queries the system health of the local
              computer (hardware, software, and network) and
              prints it.
================================================================================
#>
& "$PSScriptRoot/check-hardware.ps1"
& "$PSScriptRoot/check-software.ps1"
& "$PSScriptRoot/check-network.ps1"
exit 0 # success