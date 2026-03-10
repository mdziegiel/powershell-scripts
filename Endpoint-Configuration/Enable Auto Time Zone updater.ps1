# Ensure required services for location-based time zone + time sync
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
