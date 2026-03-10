<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2024/03/20

  Script: Remove 260 Character Path Limit.ps1

  Scope: Remove 260 Character Path Limit
 
==========================================================================
#>

New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force