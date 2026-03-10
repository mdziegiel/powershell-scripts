<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2023/12/13

  Script: FireFoxRemoval.ps1

  Scope: Removal of Firefox browser
 
==========================================================================
#>

try 
{
$appExists = $true 
$filePath = "C:\Program Files\Mozilla Firefox\uninstall\helper.exe" 
if(-not(Test-Path $filePath)) 
{ 
$filePath = "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe" 
if(-not(Test-Path $filePath)) 
{ 
$appExists = $false 
} 
}
if($appExists) 
{
Start-Process -FilePath $filePath -ArgumentList "/s" -Wait
Write-Host "Successfully uninstalled Firefox application." 
}
else
{
Write-Host "Firefox is not installed on this device.”} 
}
catch 
{ 
Write-Host $_.Execption.Message 
}