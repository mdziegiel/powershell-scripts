<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Remove-Widget
SYNOPSIS    : Removes the Windows 11 Widgets panel.
DESCRIPTION : Removes the Windows 11 Widgets / WebExperience package for
              existing users and removes the provisioned package
              for new users.
================================================================================
#>
Get-AppxPackage -AllUsers | Where-Object {$_.Name -like "*WebExperience*"} | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue

# Remove the provisioned package for new users
$AppxRemoval = Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*WebExperience*"} 
ForEach ( $App in $AppxRemoval) {
    Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName 
}