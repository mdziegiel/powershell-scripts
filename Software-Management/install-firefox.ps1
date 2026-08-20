<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : install-firefox
SYNOPSIS    : This PowerShell script installs the Mozilla Firefox browser from
              Microsoft Store.
DESCRIPTION : This PowerShell script installs the Mozilla Firefox browser from
              Microsoft Store.
================================================================================
#>
try {
	"⏳ Installing Mozilla Firefox from Microsoft Store..."
	$stopWatch = [system.diagnostics.stopwatch]::startNew()

	& winget install --id 9NZVDKPMR9RD --source msstore --accept-package-agreements --accept-source-agreements
	if ($lastExitCode -ne 0) { throw "Can't install Mozilla Firefox - maybe it's already installed" }

	[int]$elapsed = $stopWatch.Elapsed.TotalSeconds
	"✅ Mozilla Firefox installed successfully in $($elapsed)s."
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (in script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}