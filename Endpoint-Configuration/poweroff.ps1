<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : poweroff
SYNOPSIS    : This script halts the local computer immediately (needs admin
              rights).
DESCRIPTION : This script halts the local computer immediately (needs admin
              rights).
================================================================================
#>
#Requires -RunAsAdministrator
try {
	if ($IsLinux) {
		sudo shutdown
	} else {
		Stop-Computer
	}
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}