<# 
CreateShortcut.ps1
Creates Measure.lnk on Public Desktop and drops a marker file so Intune detects it.
Designed to run as SYSTEM during Autopilot ESP.
#>

$ErrorActionPreference = 'Stop'

# ====== CONFIG ======
$ShortcutName = 'Measure.lnk'
$TargetPath   = '\\hkfiles\e\Net_Admin\Software\Measure\Measure_hk.exe'
$WorkingDir   = '\\hkfiles\e\Net_Admin\Software\Measure'

$PublicDesktop = 'C:\Users\Public\Desktop'

# Intune detection marker
$MarkerDir  = 'C:\ProgramData\HK\MeasureShortcut'
$MarkerFile = Join-Path $MarkerDir 'installed.tag'
$LogFile    = Join-Path $MarkerDir 'install.log'
# ====================

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format s
    "[$stamp] $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
}

try {
    # Ensure marker/log folder exists
    New-Item -ItemType Directory -Path $MarkerDir -Force | Out-Null

    Write-Log "=== Starting Measure shortcut install ==="
    Write-Log "Running as: $env:USERNAME"
    Write-Log "ShortcutPath: $(Join-Path $PublicDesktop $ShortcutName)"
    Write-Log "TargetPath:   $TargetPath"

    $ShortcutPath = Join-Path $PublicDesktop $ShortcutName

    # Do not hard-fail if share isn't reachable during ESP; shortcut can still point to UNC
    if (-not (Test-Path $TargetPath)) {
        Write-Log "WARNING: Target not reachable right now (OK during ESP): $TargetPath"
    }

    # Remove existing shortcut if present
    if (Test-Path $ShortcutPath) {
        Remove-Item -Path $ShortcutPath -Force -ErrorAction SilentlyContinue
        Write-Log "Removed existing shortcut: $ShortcutPath"
    }

    # Create shortcut
    $wsh = New-Object -ComObject WScript.Shell
    $sc  = $wsh.CreateShortcut($ShortcutPath)
    $sc.TargetPath       = $TargetPath
    $sc.WorkingDirectory = $WorkingDir
    $sc.WindowStyle      = 1
    $sc.Save()

    Write-Log "Created shortcut: $ShortcutPath"

    # Create marker for Intune detection (ONLY after success)
    "installed $(Get-Date -Format s)" | Out-File -FilePath $MarkerFile -Force -Encoding ascii
    Write-Log "Wrote marker: $MarkerFile"

    # Extra sanity check
    if (-not (Test-Path $MarkerFile)) {
        throw "Marker file was not created for unknown reason."
    }

    Write-Log "=== SUCCESS ==="
    exit 0
}
catch {
    # Always log failures
    try {
        New-Item -ItemType Directory -Path $MarkerDir -Force | Out-Null
        Write-Log "=== FAILURE ==="
        Write-Log $_.Exception.Message
        Write-Log $_.ScriptStackTrace
    } catch {}

    exit 1
}
