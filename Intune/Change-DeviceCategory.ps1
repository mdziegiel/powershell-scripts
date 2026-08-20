<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Change-DeviceCategory
SYNOPSIS    : Sets the device category for a single Intune device using the
              Microsoft Graph beta endpoint via the Intune
              PowerShell SDK (Connect-MSGraph).
DESCRIPTION : Sets the device category for a single Intune device using the
              Microsoft Graph beta endpoint via the Intune
              PowerShell SDK (Connect-MSGraph).
================================================================================
#>
#Requires -Modules Microsoft.Graph.Authentication
Connect-MSGraph
Update-MSGraphEnvironment -SchemaVersion 'beta'

function Set-DeviceCategory {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ $_ -ne 'ADD-DEVICE-ID' })]
        [string]$DeviceID,

        [Parameter(Mandatory)]
        [ValidateScript({ $_ -ne 'ADD-THE-DEVICE-CATEGORY-ID' })]
        [string]$DeviceCategory
    )

    try {
        $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$DeviceCategory" }
        Invoke-MSGraphRequest -HttpMethod PUT -Url "deviceManagement/managedDevices/$DeviceID/deviceCategory/`$ref" -Content $body
    }
    catch {
        Write-Error "Failed to set device category: $_"
    }
}

$DeviceID = 'ADD-DEVICE-ID'
$DeviceCategory = 'ADD-THE-DEVICE-CATEGORY-ID'

Set-DeviceCategory -DeviceID $DeviceID -DeviceCategory $DeviceCategory