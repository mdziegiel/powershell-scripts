<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Change-DeviceCategorySingle
SYNOPSIS    : Sets the device category for one Intune device using the Microsoft
              Graph beta endpoint via the Intune PowerShell SDK
              (Connect-MSGraph).
DESCRIPTION : Sets the device category for one Intune device using the Microsoft
              Graph beta endpoint via the Intune PowerShell SDK
              (Connect-MSGraph).
================================================================================
#>
#Requires -Modules Microsoft.Graph.Authentication
Connect-MSGraph
Update-MSGraphEnvironment -SchemaVersion 'beta'

function Change-DeviceCategory {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceID,

        [Parameter(Mandatory)]
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

try {
    Change-DeviceCategory -DeviceID $DeviceID -DeviceCategory $DeviceCategory
}
catch {
    Write-Error "Failed to change device category: $_"
}