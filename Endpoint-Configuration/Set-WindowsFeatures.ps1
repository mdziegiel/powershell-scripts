<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Set Windows Features
SYNOPSIS    : Enables or disables Windows optional features
DESCRIPTION : Manages Windows optional features using DISM. Supports
              predefined profiles (Standard, Developer, Kiosk) or
              custom feature lists for enabling or disabling features.
              Includes logging for auditing changes
==============================================================================
#>

#Requires -Version 5.1
<#
.SYNOPSIS
    Enables or disables Windows optional features in bulk.
.DESCRIPTION
    Manages Windows optional features via DISM. Includes predefined profiles
    for common Hans Kissle endpoint configurations. Can also accept a custom
    feature list. Logs all changes for auditing.
.PARAMETER Profile
    Predefined feature set to apply. Options: Standard, Developer, Kiosk.
.PARAMETER Enable
    Array of specific feature names to enable.
.PARAMETER Disable
    Array of specific feature names to disable.
.PARAMETER OutputPath
    Path for the log file. Defaults to script directory.
.EXAMPLE
    .\Set-WindowsFeatures.ps1 -Profile Standard
    .\Set-WindowsFeatures.ps1 -Enable "TelnetClient","TFTP"
    .\Set-WindowsFeatures.ps1 -Disable "WindowsMediaPlayer","Internet-Explorer-Optional-amd64"
    .\Set-WindowsFeatures.ps1 -Profile Standard -OutputPath "C:\Logs"
#>

[CmdletBinding(DefaultParameterSetName="Profile")]
param(
    [Parameter(ParameterSetName="Profile")]
    [ValidateSet("Standard","Developer","Kiosk")]
    [string]$Profile = "Standard",

    [Parameter(ParameterSetName="Custom")]
    [string[]]$Enable,

    [Parameter(ParameterSetName="Custom")]
    [string[]]$Disable,

    [string]$OutputPath = $PSScriptRoot
)

# Requires elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$LogPath = Join-Path $OutputPath "WindowsFeatures_$Timestamp.log"
$Log = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $Entry = "$(Get-Date -Format 'HH:mm:ss') | $Message"
    $Log.Add($Entry)
    Write-Host $Entry -ForegroundColor $Color
}

# Define profiles
$Profiles = @{
    Standard = @{
        Enable  = @(
            "NetFx3"                                    # .NET Framework 3.5
        )
        Disable = @(
            "WindowsMediaPlayer",                       # Windows Media Player
            "Internet-Explorer-Optional-amd64",         # Internet Explorer
            "WorkFolders-Client",                       # Work Folders
            "Printing-XPSServices-Features"             # XPS Services
        )
    }
    Developer = @{
        Enable  = @(
            "NetFx3",                                   # .NET Framework 3.5
            "TelnetClient",                             # Telnet Client
            "TFTP",                                     # TFTP Client
            "HypervisorPlatform",                       # Hypervisor Platform
            "VirtualMachinePlatform",                   # Virtual Machine Platform
            "Microsoft-Windows-Subsystem-Linux"         # WSL
        )
        Disable = @(
            "WindowsMediaPlayer",
            "Internet-Explorer-Optional-amd64"
        )
    }
    Kiosk = @{
        Enable  = @(
            "NetFx3"
        )
        Disable = @(
            "WindowsMediaPlayer",
            "Internet-Explorer-Optional-amd64",
            "WorkFolders-Client",
            "Printing-XPSServices-Features",
            "TFTP",
            "TelnetClient",
            "Microsoft-Windows-Subsystem-Linux",
            "HypervisorPlatform",
            "VirtualMachinePlatform"
        )
    }
}

# Resolve feature lists
if ($PSCmdlet.ParameterSetName -eq "Profile") {
    $EnableList  = $Profiles[$Profile].Enable
    $DisableList = $Profiles[$Profile].Disable
    Write-Log "Applying profile: $Profile" "Cyan"
} else {
    $EnableList  = $Enable
    $DisableList = $Disable
    Write-Log "Applying custom feature list." "Cyan"
}

# Enable features
foreach ($Feature in $EnableList) {
    Write-Log "Enabling: $Feature" "Cyan"
    try {
        $Result = Enable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart -ErrorAction Stop
        $Status = if ($Result.RestartNeeded) { "Enabled (reboot required)" } else { "Enabled" }
        Write-Log "$Feature - $Status" "Green"
    }
    catch {
        Write-Log "$Feature - FAILED: $($_.Exception.Message)" "Red"
    }
}

# Disable features
foreach ($Feature in $DisableList) {
    Write-Log "Disabling: $Feature" "Cyan"
    try {
        $Result = Disable-WindowsOptionalFeature -Online -FeatureName $Feature -NoRestart -ErrorAction Stop
        $Status = if ($Result.RestartNeeded) { "Disabled (reboot required)" } else { "Disabled" }
        Write-Log "$Feature - $Status" "Green"
    }
    catch {
        Write-Log "$Feature - FAILED or not present: $($_.Exception.Message)" "Yellow"
    }
}

# Save log
$Log | Out-File -FilePath $LogPath -Encoding UTF8
Write-Log "`nComplete. Log saved to: $LogPath" "Yellow"
