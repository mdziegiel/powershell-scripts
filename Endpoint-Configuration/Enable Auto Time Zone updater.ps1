<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Configure Time Sync and Time Zone Services
SYNOPSIS    : Ensures time sync and location-based time zone services are running
DESCRIPTION : Sets required Windows services (lfsvc, tzautoupdate, w32time)
              to Automatic and starts them if needed. Repairs the Windows
              Time service if misconfigured and forces a time resync
==============================================================================
#>

Set-Service lfsvc -StartupType Automatic
Set-Service tzautoupdate -StartupType Automatic
Set-Service w32time -StartupType Automatic
 
Start-Service lfsvc -ErrorAction SilentlyContinue
Start-Service tzautoupdate -ErrorAction SilentlyContinue
Start-Service w32time -ErrorAction SilentlyContinue
 
# If time service is in a bad state, repair it (only if needed)
$source = (w32tm /query /source) 2>$null
if ($source -match "Unknown") {
    Stop-Service w32time -Force -ErrorAction SilentlyContinue
    w32tm /unregister | Out-Null
    w32tm /register   | Out-Null
    Set-Service w32time -StartupType Automatic
    Start-Service w32time -ErrorAction SilentlyContinue
}
 
# Force a resync
w32tm /resync /force | Out-Null
