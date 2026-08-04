<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Canon Printer Removal
SYNOPSIS    : Removes Canon printers
DESCRIPTION : Identifies and removes Canon printer objects

ORGANIZATION: 
==============================================================================
#>

$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.222*"}
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.248*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.9*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.11*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.8*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*10.103.1.46*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
