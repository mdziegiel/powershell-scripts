<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2024/08/09

  Script: ScreenConnectRemoval.ps1

  Scope: Removal of ScreenConnect
 
==========================================================================
#>

Get-Package -Name "ScreenConnect Client (d015d7e65cecdce1)" | Uninstall-Package