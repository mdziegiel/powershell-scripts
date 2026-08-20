<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : install-git-for-windows
SYNOPSIS    : This PowerShell script installs Git for Windows.
DESCRIPTION : This PowerShell script installs Git for Windows.
================================================================================
#>
try {
	"Installing Git for Windows, please wait..."

	& winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
	if ($lastExitCode -ne 0) { throw "'winget install' failed" }

	"Git for Windows installed successfully."
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}