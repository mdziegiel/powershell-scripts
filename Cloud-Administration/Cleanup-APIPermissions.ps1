<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Cleanup-APIPermissions
SYNOPSIS    : To enhance your tenant's security posture, it's crucial to
              regularly review the API permissions requested by
              SPFx solutions and compare them with those granted
              to the ”SharePoint Online Client Extensibility Web…
DESCRIPTION : To enhance your tenant's security posture, it's crucial to
              regularly review the API permissions requested by
              SPFx solutions and compare them with those granted
              to the ”SharePoint Online Client Extensibility Web…
================================================================================
#>
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns
Param(
    [Parameter (Mandatory = $true)]
    [String] $tenantName = "your-tenant-name",
    [Parameter (Mandatory = $true)]
    [String] $appCatalogSiteName = "appcatalog",
    [Parameter (Mandatory = $true)]
    [switch]$WhatIf = $false,
    [switch]$Interactive = $false
)
#####################################
# The following PS modules must be added to the Azure Runbook
# Microsoft.Graph.Authentication
# Microsoft.Graph.Applications
# Microsoft.Graph.Identity.SignIns
#
#####################################
$Global:SPOAppName = "SharePoint Online Client Extensibility Web Application Principal"

function Get-GrantedAPIPermissions {
    param(
    )
    Try {

        #####################################
        # Get API Permissions for SharePoint Online Client Extensibility Web Application Principal
        #####################################

        $servicePrincipal = Get-MgServicePrincipal -Filter  "DisplayName eq '$($Global:SPOAppName)'"

        #Delegated permission grants authorizing this service principal to access an API on behalf of a signed-in user.
        $permissions = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $servicePrincipal.Id

        $permissionsDelegated = $permissions | ForEach-Object {
            # retrieve delegated permissions of the resource service principal
            $resource = Get-MgServicePrincipal -ServicePrincipalId $_.ResourceId

            [PSCustomObject]@{
                Scope      = $_.Scope
                ResourceId = $_.ResourceId
                Resource   = $resource.DisplayName
                AllUsers   = $_.ConsentType -eq "AllPrincipals"
            }
        }

        #this is NOT supported for the "SharePoint Online Client Extensibility Web Application Principal". Result should always be empty
        $permissionsApplication = get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id
        If ($null -ne $permissionsApplication) {
            Write-Warning ("Assigning Application permissions to the 'SharePoint Online Client Extensibility Web Application Principal' is NOT SUPPORTED") -ForegroundColor Red
        }

        @{
            delegated   = $permissionsDelegated
            application = $permissionsApplication
        }

    }
    Catch {
        Write-Error "Error downloading API Permissions information: $($_.Exception.Message)"
    }
}
function Revoke-SelectedAPIPermissions {
    param(
        [System.Object] $delegatedPermissionsByScopeUsage ,
        [bool] $whatIf
    )
    $servicePrincipal = Get-MgServicePrincipal -Filter  "DisplayName eq '$($Global:SPOAppName)'"

    $permissions = Get-MgOauth2PermissionGrant -Filter "ClientId eq '$($servicePrincipal.Id)' and ConsentType eq 'AllPrincipals'"

    $delegatedPermissionsByScopeUsage | Group-Object -Property ResourceId  | ForEach-Object {
        $Grouped = $_
        
        $OAuth2PermissionGrant = $permissions | Where-Object { $_.ResourceId -eq $Grouped.Name }
        $OAuth2PermissionGrantId = $OAuth2PermissionGrant.id
       
        $required = ($Grouped.Group | Where-Object { $null -ne $_.spfxName }).Scope
        $notRequired = ($Grouped.Group | Where-Object { $null -eq $_.spfxName }).Scope

        if ( $null -ne $required -and $null -ne $notRequired) {
            "Updating API Permissions for $( $Grouped.Group[0].Resource). Removing $( $notRequired -join ', ')"
            $params = @{
                Scope = $required -join " "
            }
            if ($whatIf) {
                "*** Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $OAuth2PermissionGrantId -BodyParameter `$params"
                "*** With parameters:"
                $params | ConvertTo-Json 
            }
            else {
                Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $OAuth2PermissionGrantId -BodyParameter $params
            }
        }
        elseif ( $null -eq $required -and $null -ne $notRequired) {
            "Deleting all API Permissions for $( $Grouped.Group[0].Resource)."
            if ($whatIf) {
                "*** Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $OAuth2PermissionGrantId"
            }
            else {
                Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $OAuth2PermissionGrantId
            
            }
        }
    }
}
function Get-SPFxAPIPermissions {
    param(
        [string]$domainName,
        [string]$appcatalog
    )

    <#
    Function Get-SitesWithAppCatalog returns url part after /sites/ for each site with an app catalog
    #>
    function Get-SitesWithAppCatalog {
        param(
            [string]$domainName,
            [string]$appcatalog
        )

        $listTitle = "Site Collection App Catalogs"
        $expandQuery = '$expand=fields($select=id,sitecollectionurl)'
        $listItems = Invoke-MgGraphRequest -Method GET `
            "https://graph.microsoft.com/v1.0/sites/$($domainName):/sites/$($appcatalog):/lists/$($listTitle)/items?$expandQuery"

        $siteAppCatalogs = $listItems.value | ForEach-Object { 
            $url = $_.fields.SiteCollectionUrl
            $res = $url -match "sites/(?<siteName>.*)"
            if ($res) {
                $matches["siteName"]
            }
        }
        $siteAppCatalogs += $appcatalog 

        $siteAppCatalogs
    }
    function Get-SPFSolutionsAndAPIPermissions {
        param(
            [string]$domainName,
            [string[]]$siteUrls
        )

        $siteUrls | ForEach-Object {

            $siteUrl = $_
            $ErrorActionPreference = 'SilentlyContinue'

            $listTitle = "Apps for SharePoint"
            $expandQuery = '$expand=fields($select=FileLeafRef, AppVersion,WebApiPermissionScopesNote,Title)'
            $listItems = Invoke-MgGraphRequest -Method GET `
                "https://graph.microsoft.com/v1.0/sites/$($domainName):/sites/$($siteUrl):/lists/$($listTitle)/items?$expandQuery"
            
            $listItems.value | ForEach-Object { 
                if ($null -ne $_ -and $null -ne $_.fields) {
                    [PSCustomObject]@{
                        siteURL        = "https://$domainName/sites/$siteUrl"
                        fileName       = $_.fields.FileLeafRef
                        version        = "v$($_.fields.AppVersion)"
                        apiPermissions = $_.fields.WebApiPermissionScopesNote
                        title          = $_.fields.Title
                        Error          = ""
                    }
                }
                else {
                    [PSCustomObject]@{
                        siteURL        = "https://$domainName/sites/$siteUrl"
                        fileName       = ""
                        version        = ""
                        apiPermissions = ""
                        title          = ""
                        Error          = "Site not found or access denied."
                    }
                }
            }
        }
       
    }

    $siteUrls = Get-SitesWithAppCatalog -domainName $domainName -appcatalog $appcatalog
    Get-SPFSolutionsAndAPIPermissions -domainName $domainName -siteUrls $siteUrls
}
function Get-Usage {
    param(
        [System.Object]$apiPermissionsGranted, 
        [System.Object]$spfxPermissionsRequested 
    )
    
    function Get-SpfxPermissionsByScope {
        param(
            [System.Object]$spfxPermissions
        )
        $arr = @()
        $spfxPermissions | ForEach-Object {
            $fileName = $_.fileName
            $title = $_.title
            $_.apiPermissions -split "; " |  Sort-Object | Get-Unique | ForEach-Object { 
                $resource, $scope = $_.Trim() -split ", " 
                $item = $arr | Where-Object { $_.Resource -eq $resource -and $_.Scope -eq $scope }
                if ($null -eq $item) {
                    $arr += [PSCustomObject]@{
                        Resource = $resource
                        Scope    = $scope
                        fileName = @($fileName)
                        title    = @($title)
                    }
                }
                else {
                    If ($item.fileName -notcontains $fileName) {
                        $item.fileName += $fileName
                        $item.title += $title
                    }

                }
            }
        }
        $arr
    }
    function Get-GrantedPermissionsByScope {
        param(
            [System.Object] $apiPermissionsDelegated 
        )
        $apiPermissionsDelegated | ForEach-Object {
            $item = $_
            $_.Scope -split " " | ForEach-Object {
                [PSCustomObject]@{
                    Resource   = $item.Resource
                    ResourceId = $item.ResourceId
                    Scope      = $_
                    AllUsers   = $item.AllUsers
                }
            }
            
        }
    }
    function Join-Arrays {
        param (
            [Parameter(Mandatory = $true)] [array]$Left,
            [Parameter(Mandatory = $true)] [array]$Right
        )

        foreach ($leftItem in $Left) {
            $match = $Right | Where-Object { $leftItem.Resource -eq $_.Resource -and $leftItem.Scope -eq $_.Scope }
            [PSCustomObject]@{
                Resource     = $leftItem.Resource
                ResourceId   = $leftItem.ResourceId
                Scope        = $leftItem.Scope
                AllUsers     = $leftItem.AllUsers
                spfxName     = $match.fileName
                SolutionName = $match.title
            }
        }
    }


    # | Scope | Permission | array(fileName) | array(title) |
    $spfxPermissionsByScope = Get-SpfxPermissionsByScope -spfxPermissions ( $spfxPermissionsRequested | Where-Object { $null -ne $_.apiPermissions } )
    $delegatedPermissionsByScope = Get-GrantedPermissionsByScope -apiPermissionsDelegated $apiPermissionsGranted.delegated

    $delegatedPermissionsByScopeUsage = Join-Arrays -Left $delegatedPermissionsByScope -Right $spfxPermissionsByScope

    @{
        spfxPermissionsByScope           = $spfxPermissionsByScope
        delegatedPermissionsByScopeUsage = $delegatedPermissionsByScopeUsage
    }
}

function Invoke-CleanupAPIPermissions {
    Param(
        [Parameter (Mandatory = $true)]
        [string]$tenantName,
        [string] $appCatalogSiteName ,
        [bool] $whatIf ,
        [bool]$interactive = $false
    )

    If ($interactive) {
        Connect-MgGraph -Scopes "Sites.Selected", "Application.Read.All", "DelegatedPermissionGrant.ReadWrite.All"
    }
    else {
        #####################################
        # If MSI is used, two environment variables MSI_ENDPOINT and MSI_SECRET are available
        #####################################
        if ($null -eq $env:MSI_ENDPOINT) {
            Write-Error "To execute this script, please enable Managed Identity for this automation, and grant all required API permissions"
            return
        }
  
        Connect-MgGraph  -Identity -NoWelcome
    }
    "Connected"
        
    # Get API permissions granted to the SPO service principal
    $apiPermissionsGranted = Get-GrantedAPIPermissions 
    # Get API permissions requested by SPFx solutions 
    # Important: if the Managed Identity doesn't have access to a site (e.g. site-level app catalog, the results will be incomplete)
    $spfxPermissionsRequested = Get-SPFxAPIPermissions -domainName "$tenantName.sharepoint.com" -appcatalog $appCatalogSiteName
    # compare assigned and requested API permissions    
    $apiUsageInfo = Get-Usage -apiPermissionsGranted $apiPermissionsGranted  -spfxPermissionsRequested $spfxPermissionsRequested 
    
    # remove unused API permissions assignments
    Revoke-SelectedAPIPermissions -delegatedPermissionsByScopeUsage $apiUsageInfo.delegatedPermissionsByScopeUsage -whatIf $whatIf

}
# Main runbook content 
Invoke-CleanupAPIPermissions -tenantName $tenantName -appCatalogSiteName $appCatalogSiteName -whatIf $WhatIf.IsPresent -interactive $Interactive.IsPresent