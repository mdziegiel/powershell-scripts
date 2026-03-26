<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Set Wallpaper and Lock Screen
SYNOPSIS    : Sets lock screen and schedules desktop wallpaper deployment
DESCRIPTION : Copies wallpaper.jpg to a local path, applies it as the
              device-level lock screen, and schedules the desktop
              background to be configured at the next user logon
==============================================================================
#>

$ErrorActionPreference = "Stop"

$RootPath  = "C:\ProgramData\HK\Wallpaper"
$ImageName = "wallpaper.jpg"
$ImageDest = Join-Path $RootPath $ImageName

$UserScriptName = "Set-DesktopBackground-User.ps1"
$UserScriptDest = Join-Path $RootPath $UserScriptName

$MarkerFile = Join-Path $RootPath "WallpaperApplied.marker"

# Ensure folder exists
if (-not (Test-Path $RootPath)) { New-Item -Path $RootPath -ItemType Directory -Force | Out-Null }

# Copy image from package to local folder
$srcImage = Join-Path $PSScriptRoot $ImageName
if (-not (Test-Path $srcImage)) { throw "Missing image next to script: $srcImage" }
Copy-Item $srcImage -Destination $ImageDest -Force

# Copy the user-context script locally
$srcUserScript = Join-Path $PSScriptRoot $UserScriptName
if (-not (Test-Path $srcUserScript)) { throw "Missing user script next to installer: $srcUserScript" }
Copy-Item $srcUserScript -Destination $UserScriptDest -Force

# Set LOCK SCREEN image (device-level)
$lockKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
if (-not (Test-Path $lockKey)) { New-Item -Path $lockKey -Force | Out-Null }
New-ItemProperty -Path $lockKey -Name "LockScreenImage" -Value $ImageDest -PropertyType String -Force | Out-Null

# Stage DESKTOP BACKGROUND set at next logon (runs once for the next user who logs in)
$runOnceKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
$cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$UserScriptDest`""
New-ItemProperty -Path $runOnceKey -Name "HK-SetDesktopBackground" -Value $cmd -PropertyType String -Force | Out-Null

# Marker for Intune detection
New-Item -Path $MarkerFile -ItemType File -Force | Out-Null

