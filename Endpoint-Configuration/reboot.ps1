<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : reboot
SYNOPSIS    : This PowerShell script reboots the local computer immediately
              (needs admin rights).
DESCRIPTION : This PowerShell script reboots the local computer immediately
              (needs admin rights).
================================================================================
#>
#Requires -RunAsAdministrator
try {
	if ($IsLinux) {
		& sudo reboot
	} else {
		Restart-Computer
	}
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}