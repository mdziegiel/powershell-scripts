<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Canon Printer Removal
SYNOPSIS    : Removes Canon printers
DESCRIPTION : Identifies and removes Canon printer objects

ORGANIZATION: 
==============================================================================
#>

$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"}
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
$PrinterObj = Get-Printer | where {$_.portname -like "*IPAddess*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }
