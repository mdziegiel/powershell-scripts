<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2024/06/10

  Script: RemoveRegistryKeysl.ps1

  Scope: Removal of Registry Keys WindowsUpdate WindowsUpdate\AU
 
==========================================================================
#>
Remove-Item -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate -Force -Recurse