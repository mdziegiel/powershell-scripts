<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Remove uniFLOW Print Server
SYNOPSIS    : Removes gidbos89 printer connections
DESCRIPTION : Finds and removes any printers matching gidbos89 from
              the local system

ORGANIZATION: GID
==============================================================================
#>

$PrinterObj = Get-Printer | where {$_.Name -like "*gidbos89*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
