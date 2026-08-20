<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get-CleanUpDiskDetection
SYNOPSIS    : Checks free space on the C: drive
DESCRIPTION : Checks free space on the C: drive. If free space is below the
              threshold (15 GB), exits with code 1 to trigger
              remediation.
================================================================================
#>
$storageThreshold = 15

$utilization = (Get-PSDrive | Where-Object { $_.Name -eq "C" }).Free

if (($storageThreshold * 1GB) -lt $utilization) {
    exit 0
} else {
    exit 1
}