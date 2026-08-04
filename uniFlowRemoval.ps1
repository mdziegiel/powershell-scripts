<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Remove uniFLOW Print Server
SYNOPSIS    : Removes uniFLOW printer connections
DESCRIPTION : Finds and removes any printers matching "servername" from
              the local system

ORGANIZATION: 
==============================================================================
#>

$PrinterObj = Get-Printer | where {$_.Name -like "*servername*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
