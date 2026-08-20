<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : install-ssh-client
SYNOPSIS    : This PowerShell script installs a SSH client (needs admin rights).
DESCRIPTION : This PowerShell script installs a SSH client (needs admin rights).
================================================================================
#>
try {
	$StopWatch = [system.diagnostics.stopwatch]::startNew()

	if ($IsLinux) {
		& sudo apt install openssh-client
	} else {
		Add-WindowsCapability -Online -Name OpenSSH.Client*
	}

	[int]$Elapsed = $StopWatch.Elapsed.TotalSeconds
	"✅ installed SSH client in $Elapsed sec"
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}