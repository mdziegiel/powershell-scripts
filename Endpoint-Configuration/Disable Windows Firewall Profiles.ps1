<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Disable Windows Firewall Profiles
SYNOPSIS    : Disables Windows Defender Firewall on the local device
DESCRIPTION : Disables Domain, Private, and Public firewall profiles.
              Designed for local execution or Intune deployment
==============================================================================
#>

try {
    Write-Host "Disabling Windows Firewall for all profiles..."

    Set-NetFirewallProfile -Profile Domain,Private,Public -Enabled False

    $status = Get-NetFirewallProfile | Select-Object Name, Enabled

    Write-Host "Current Firewall Status:"
    $status | ForEach-Object {
        Write-Host "$($_.Name) : Enabled = $($_.Enabled)"
    }

    exit 0
}
catch {
    Write-Host "Error disabling firewall: $_"
    exit 1
}
