#Requires -Version 5.1
<#
.SYNOPSIS
    Configures Windows power settings for corporate workstations.
.DESCRIPTION
    Sets sleep, hibernate, display timeout, and fast startup settings
    via powercfg. Designed for Hans Kissle domain-joined workstations
    to ensure consistent power behavior across the environment.
.PARAMETER Profile
    Power profile to apply. Options: Workstation, Laptop, Kiosk.
    Workstation: longer timeouts, hibernate disabled.
    Laptop: balanced timeouts, hibernate enabled.
    Kiosk: display never off, sleep never, hibernate disabled.
.EXAMPLE
    .\Set-PowerConfiguration.ps1
    .\Set-PowerConfiguration.ps1 -Profile Laptop
    .\Set-PowerConfiguration.ps1 -Profile Kiosk
#>

[CmdletBinding()]
param(
    [ValidateSet("Workstation","Laptop","Kiosk")]
    [string]$Profile = "Workstation"
)

# Requires elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Write-Host "Applying power configuration profile: $Profile" -ForegroundColor Cyan

# Set balanced power plan as base
powercfg /setactive SCHEME_BALANCED | Out-Null

switch ($Profile) {
    "Workstation" {
        # AC power settings (plugged in)
        powercfg /change monitor-timeout-ac 30        # Display off after 30 min
        powercfg /change standby-timeout-ac 0         # Never sleep
        powercfg /change hibernate-timeout-ac 0       # Never hibernate
        powercfg /change disk-timeout-ac 0            # Never spin down disk

        # DC power settings (battery - not applicable but set anyway)
        powercfg /change monitor-timeout-dc 10
        powercfg /change standby-timeout-dc 30
        powercfg /change hibernate-timeout-dc 60

        # Disable hibernate entirely on desktops
        powercfg /hibernate off
        Write-Host "Workstation profile applied: Display off 30min, sleep/hibernate disabled." -ForegroundColor Green
    }
    "Laptop" {
        # AC
        powercfg /change monitor-timeout-ac 15
        powercfg /change standby-timeout-ac 30
        powercfg /change hibernate-timeout-ac 0

        # DC
        powercfg /change monitor-timeout-dc 5
        powercfg /change standby-timeout-dc 15
        powercfg /change hibernate-timeout-dc 30

        # Enable hibernate for laptops
        powercfg /hibernate on
        Write-Host "Laptop profile applied: Balanced timeouts, hibernate enabled." -ForegroundColor Green
    }
    "Kiosk" {
        # Never turn off display, never sleep, never hibernate
        powercfg /change monitor-timeout-ac 0
        powercfg /change standby-timeout-ac 0
        powercfg /change hibernate-timeout-ac 0
        powercfg /change monitor-timeout-dc 0
        powercfg /change standby-timeout-dc 0
        powercfg /change hibernate-timeout-dc 0
        powercfg /hibernate off

        # Disable fast startup (can cause issues on kiosks)
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" `
            -Name "HiberbootEnabled" -Value 0 -Type DWord -Force

        Write-Host "Kiosk profile applied: Display always on, sleep/hibernate disabled." -ForegroundColor Green
    }
}

# Disable fast startup globally (recommended for domain machines)
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" `
    -Name "HiberbootEnabled" -Value 0 -Type DWord -Force
Write-Host "Fast startup disabled." -ForegroundColor Green

Write-Host "`nPower configuration complete." -ForegroundColor Yellow
