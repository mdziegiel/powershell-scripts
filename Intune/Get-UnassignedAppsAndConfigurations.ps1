<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get-UnassignedAppsAndConfigurations
SYNOPSIS    : Queries Microsoft Graph for all mobile apps with their assignments
              expanded and returns those that have no assignments
              configured.
DESCRIPTION : Queries Microsoft Graph for all mobile apps with their assignments
              expanded and returns those that have no assignments
              configured.
================================================================================
#>
#Requires -Modules Microsoft.Graph.Authentication
function Connect-MgGraphIfNeeded {
    $context = Get-MgContext
    if (-not $context) {
        Connect-MgGraph -Scopes "DeviceManagementApps.Read.All" -NoWelcome
    }
}

########################################################################################
####################################     Start      ####################################
########################################################################################
# Auth
Connect-MgGraphIfNeeded

$uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$" + "expand=assignments"
$devices = @()
try {
    while ($uri) {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
        $devices += $response.value
        $uri = $response.'@odata.nextLink'
    }
}
catch {
    Write-Error "Failed to retrieve apps: $_"
    throw
}