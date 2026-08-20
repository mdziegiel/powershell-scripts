<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : install-chrome
SYNOPSIS    : This PowerShell script installs the Google Chrome browser from
              WinGet.
DESCRIPTION : This PowerShell script installs the Google Chrome browser from
              WinGet.
================================================================================
#>
try {
	"⏳ Installing Google Chrome from WinGet..."
	$stopWatch = [system.diagnostics.stopwatch]::startNew()

	& winget install --id Google.Chrome --accept-package-agreements --accept-source-agreements
	if ($lastExitCode -ne 0) { throw "Can't install Google Chrome - maybe it's already installed" }

	[int]$elapsed = $stopWatch.Elapsed.TotalSeconds
	"✅ Google Chrome installed successfully in $($elapsed)s."
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}