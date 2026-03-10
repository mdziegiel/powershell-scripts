<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2023/12/13

  Script: BostonuniFlowRemoval.ps1

  Scope: Removes gidbos89 Print Server
 
==========================================================================
#>

$PrinterObj = Get-Printer | where {$_.Name -like "*gidbos89*"} 
if ($PrinterObj -ne $null) {remove-printer -inputobject $PrinterObj }