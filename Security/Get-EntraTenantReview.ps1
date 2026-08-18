<#
AUTHOR      : Michael Dziegiel
SCRIPT      : Get Entra Tenant Review Report
SYNOPSIS    : Exports Entra tenant review reports
DESCRIPTION : Exports Entra user, admin, policy, and application reports.
NOTES       : Manual run script. Output saves to C:\Reporting.
              Requires Microsoft.Graph PowerShell and appropriate Graph permissions.
              Creates CSV files and one Excel workbook when ImportExcel is available.
#>

$ErrorActionPreference = "Stop"

# =====================================================
# Settings
# =====================================================

$ReportPath = "C:\Reporting"
$RunDate    = Get-Date -Format "yyyy-MM-dd"
$ExcelFile  = Join-Path $ReportPath "EntraTenantReview_$RunDate.xlsx"

New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null

# =====================================================
# Helper Functions
# =====================================================

function ConvertTo-FlatString {
    param([object] $Value)

    if ($null -eq $Value) { return "" }
    if ($Value -is [string]) { return $Value }

    if ($Value -is [System.Collections.IEnumerable]) {
        try {
            return (($Value | ForEach-Object { $_.ToString() }) -join "; ")
        }
        catch {
            return ($Value | ConvertTo-Json -Compress -Depth 8)
        }
    }

    return $Value.ToString()
}

function Ensure-ReportData {
    param([object[]] $Data)

    if ($null -eq $Data -or $Data.Count -eq 0) {
        return @([PSCustomObject]@{ Status = "No data returned or access not available" })
    }

    return $Data
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Data,
        [Parameter(Mandatory = $true)] [string] $FileName
    )

    $Path = Join-Path $ReportPath $FileName
    $SafeData = Ensure-ReportData -Data $Data
    $SafeData | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    return $Path
}

function Add-ExcelSheet {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Data,
        [Parameter(Mandatory = $true)] [string] $WorksheetName
    )

    if ($script:ExcelAvailable -eq $true) {
        $SafeData = Ensure-ReportData -Data $Data
        $TableName = ($WorksheetName -replace '[^a-zA-Z0-9]', '')

        $SafeData | Export-Excel -Path $ExcelFile `
                                 -WorksheetName $WorksheetName `
                                 -AutoSize `
                                 -FreezeTopRow `
                                 -BoldTopRow `
                                 -TableName $TableName `
                                 -TableStyle Medium2 `
                                 -Append
    }
}

function Set-ExcelTabColors {
    param([string] $Path)

    if ($script:ExcelAvailable -ne $true) { return }
    if (-not (Test-Path $Path)) { return }

    $TabColors = @{
        "Tenant Summary"  = "DarkBlue"
        "Active Users"    = "Green"
        "Disabled Users"  = "Red"
        "Licensed Users"  = "LightGreen"
        "Guest Users"     = "Yellow"
        "Directory Roles" = "Orange"
        "Global Admins"   = "DarkRed"
        "CA Policies"     = "Purple"
        "CA Exclusions"   = "Violet"
        "Named Locations" = "Teal"
        "Enterprise Apps" = "Blue"
        "SSO Apps"        = "LightBlue"
        "Privileged & Excluded"     = "DarkOrange"
        "Report Index"    = "Gray"
    }

    try {
        $Package = Open-ExcelPackage -Path $Path

        foreach ($SheetName in $TabColors.Keys) {
            $Worksheet = $Package.Workbook.Worksheets[$SheetName]
            if ($null -ne $Worksheet) {
                $Worksheet.TabColor = [System.Drawing.Color]::FromName($TabColors[$SheetName])
            }
        }

        Close-ExcelPackage $Package
    }
    catch {
        Write-Warning "Workbook was created, but tab colors could not be applied. $($_.Exception.Message)"
    }
}

# =====================================================
# Connect to Microsoft Graph
# =====================================================

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "User.Read.All",`
    "Directory.Read.All",`
    "Application.Read.All",`
    "RoleManagement.Read.Directory",`
    "Policy.Read.All" `
    -NoWelcome

# =====================================================
# Excel Support
# =====================================================

$script:ExcelAvailable = $false

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module not found. Installing for current user..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module ImportExcel -ErrorAction Stop
    $script:ExcelAvailable = $true

    if (Test-Path $ExcelFile) {
        Remove-Item $ExcelFile -Force
    }

    Write-Host "Excel workbook enabled: $ExcelFile" -ForegroundColor Green
}
catch {
    Write-Warning "ImportExcel is not available. CSV files will still be created, but Excel workbook will be skipped."
}

# =====================================================
# Tenant Summary
# =====================================================

Write-Host "Collecting Tenant Summary..." -ForegroundColor Cyan

$Org = Get-MgOrganization | Select-Object -First 1

$TenantSummary = @([PSCustomObject]@{
    TenantName        = $Org.DisplayName
    TenantId          = $Org.Id
    TenantType        = $Org.TenantType
    CountryLetterCode = $Org.CountryLetterCode
    VerifiedDomains   = (($Org.VerifiedDomains | ForEach-Object { $_.Name }) -join "; ")
    ReportRunDate     = Get-Date
    ReportRunBy       = (Get-MgContext).Account
})

$TenantSummaryPath = Export-ReportCsv -Data $TenantSummary -FileName "TenantSummary.csv"
Add-ExcelSheet -Data $TenantSummary -WorksheetName "Tenant Summary"

# =====================================================
# Users
# =====================================================

Write-Host "Collecting Users..." -ForegroundColor Cyan

$Users = Get-MgUser -All -Property DisplayName,UserPrincipalName,Mail,Department,JobTitle,AccountEnabled,CreatedDateTime,UserType,Description,AssignedLicenses |
Select-Object DisplayName,
              UserPrincipalName,
              Mail,
              Department,
              JobTitle,
              UserType,
              Description,
              AccountEnabled,
              CreatedDateTime,
              @{Name="AssignedLicenseCount";Expression={ if ($_.AssignedLicenses) { $_.AssignedLicenses.Count } else { 0 } }}

$ActiveUsers   = $Users | Where-Object { $_.AccountEnabled -eq $true } | Sort-Object DisplayName
$DisabledUsers = $Users | Where-Object { $_.AccountEnabled -eq $false } | Sort-Object DisplayName
$LicensedUsers = $Users | Where-Object { $_.AssignedLicenseCount -gt 0 } | Sort-Object DisplayName
$GuestUsers    = $Users | Where-Object { $_.UserType -eq "Guest" } | Sort-Object DisplayName

$ActiveUsersPath   = Export-ReportCsv -Data $ActiveUsers -FileName "ActiveUsers.csv"
$DisabledUsersPath = Export-ReportCsv -Data $DisabledUsers -FileName "DisabledUsers.csv"
$LicensedUsersPath = Export-ReportCsv -Data $LicensedUsers -FileName "LicensedUsers.csv"
$GuestUsersPath    = Export-ReportCsv -Data $GuestUsers -FileName "GuestUsers.csv"

Add-ExcelSheet -Data $ActiveUsers -WorksheetName "Active Users"
Add-ExcelSheet -Data $DisabledUsers -WorksheetName "Disabled Users"
Add-ExcelSheet -Data $LicensedUsers -WorksheetName "Licensed Users"
Add-ExcelSheet -Data $GuestUsers -WorksheetName "Guest Users"

# =====================================================
# Directory Role Assignments and Global Admins
# =====================================================

Write-Host "Collecting Directory Roles and Global Administrators..." -ForegroundColor Cyan

$RoleAssignments = @()
$DirectoryRoles = Get-MgDirectoryRole

foreach ($Role in $DirectoryRoles) {
    try {
        $Members = Get-MgDirectoryRoleMember -DirectoryRoleId $Role.Id
    }
    catch {
        Write-Warning "Could not retrieve members for role $($Role.DisplayName)."
        $Members = @()
    }

    foreach ($Member in $Members) {
        try {
            $User = Get-MgUser -UserId $Member.Id -Property DisplayName,UserPrincipalName,Mail,Description,AccountEnabled
        }
        catch {
            $User = $null
        }

        $RoleAssignments += [PSCustomObject]@{
            RoleName          = $Role.DisplayName
            MemberObjectId    = $Member.Id
            DisplayName       = if ($User) { $User.DisplayName } else { "Unable to resolve as user" }
            UserPrincipalName = if ($User) { $User.UserPrincipalName } else { "" }
            Mail              = if ($User) { $User.Mail } else { "" }
            Description       = if ($User) { $User.Description } else { "" }
            AccountEnabled    = if ($User) { $User.AccountEnabled } else { "" }
        }
    }
}

$RoleAssignments = $RoleAssignments | Sort-Object RoleName, DisplayName
$GlobalAdmins    = $RoleAssignments | Where-Object { $_.RoleName -eq "Global Administrator" } | Sort-Object DisplayName

$DirectoryRolesPath = Export-ReportCsv -Data $RoleAssignments -FileName "DirectoryRoleAssignments.csv"
$GlobalAdminsPath   = Export-ReportCsv -Data $GlobalAdmins -FileName "GlobalAdmins.csv"

Add-ExcelSheet -Data $RoleAssignments -WorksheetName "Directory Roles"
Add-ExcelSheet -Data $GlobalAdmins -WorksheetName "Global Admins"

# =====================================================
# Conditional Access Policies and Exclusions
# =====================================================

Write-Host "Collecting Conditional Access Policies..." -ForegroundColor Cyan

$CAPolicies = @()
$CAExclusions = @()

try {
    $RawCAPolicies = Get-MgIdentityConditionalAccessPolicy -All

    foreach ($Policy in $RawCAPolicies) {
        $CAPolicies += [PSCustomObject]@{
            DisplayName         = $Policy.DisplayName
            State               = $Policy.State
            PolicyId            = $Policy.Id
            CreatedDateTime     = $Policy.CreatedDateTime
            ModifiedDateTime    = $Policy.ModifiedDateTime
            IncludeUsers        = ConvertTo-FlatString $Policy.Conditions.Users.IncludeUsers
            ExcludeUsers        = ConvertTo-FlatString $Policy.Conditions.Users.ExcludeUsers
            IncludeGroups       = ConvertTo-FlatString $Policy.Conditions.Users.IncludeGroups
            ExcludeGroups       = ConvertTo-FlatString $Policy.Conditions.Users.ExcludeGroups
            IncludeRoles        = ConvertTo-FlatString $Policy.Conditions.Users.IncludeRoles
            ExcludeRoles        = ConvertTo-FlatString $Policy.Conditions.Users.ExcludeRoles
            IncludeApplications = ConvertTo-FlatString $Policy.Conditions.Applications.IncludeApplications
            ExcludeApplications = ConvertTo-FlatString $Policy.Conditions.Applications.ExcludeApplications
            IncludeLocations    = ConvertTo-FlatString $Policy.Conditions.Locations.IncludeLocations
            ExcludeLocations    = ConvertTo-FlatString $Policy.Conditions.Locations.ExcludeLocations
            GrantControls       = ConvertTo-FlatString $Policy.GrantControls.BuiltInControls
            SessionControls     = ($Policy.SessionControls | ConvertTo-Json -Compress -Depth 8)
        }

        foreach ($ExcludedUserId in @($Policy.Conditions.Users.ExcludeUsers)) {
            if ($ExcludedUserId -and $ExcludedUserId -ne "GuestsOrExternalUsers") {
                try {
                    $ResolvedUser = Get-MgUser -UserId $ExcludedUserId -Property DisplayName,UserPrincipalName,Mail,AccountEnabled,Description
                }
                catch {
                    $ResolvedUser = $null
                }

                $CAExclusions += [PSCustomObject]@{
                    PolicyName        = $Policy.DisplayName
                    PolicyState       = $Policy.State
                    ExcludedUserId    = $ExcludedUserId
                    DisplayName       = if ($ResolvedUser) { $ResolvedUser.DisplayName } else { "Unable to resolve" }
                    UserPrincipalName = if ($ResolvedUser) { $ResolvedUser.UserPrincipalName } else { "" }
                    Mail              = if ($ResolvedUser) { $ResolvedUser.Mail } else { "" }
                    AccountEnabled    = if ($ResolvedUser) { $ResolvedUser.AccountEnabled } else { "" }
                    Description       = if ($ResolvedUser) { $ResolvedUser.Description } else { "" }
                }
            }
        }
    }
}
catch {
    Write-Warning "Could not retrieve Conditional Access policies. Confirm Policy.Read.All permission is available."
}

$CAPolicies = $CAPolicies | Sort-Object DisplayName
$CAExclusions = $CAExclusions | Sort-Object PolicyName, DisplayName

$CAPoliciesPath = Export-ReportCsv -Data $CAPolicies -FileName "ConditionalAccessPolicies.csv"
$CAExclusionsPath = Export-ReportCsv -Data $CAExclusions -FileName "ConditionalAccessExclusions.csv"

Add-ExcelSheet -Data $CAPolicies -WorksheetName "CA Policies"
Add-ExcelSheet -Data $CAExclusions -WorksheetName "CA Exclusions"

# =====================================================
# Named Locations
# =====================================================

Write-Host "Collecting Named Locations..." -ForegroundColor Cyan

$NamedLocations = @()

try {
    $NamedLocations = Get-MgIdentityConditionalAccessNamedLocation -All |
    ForEach-Object {
        [PSCustomObject]@{
            DisplayName      = $_.DisplayName
            Id               = $_.Id
            CreatedDateTime  = $_.CreatedDateTime
            ModifiedDateTime = $_.ModifiedDateTime
            Type             = $_.AdditionalProperties.'@odata.type'
            Details          = ($_.AdditionalProperties | ConvertTo-Json -Compress -Depth 10)
        }
    } | Sort-Object DisplayName
}
catch {
    Write-Warning "Could not retrieve Named Locations. Confirm Policy.Read.All permission is available."
}

$NamedLocationsPath = Export-ReportCsv -Data $NamedLocations -FileName "NamedLocations.csv"
Add-ExcelSheet -Data $NamedLocations -WorksheetName "Named Locations"

# =====================================================
# Enterprise Applications and SSO Applications
# =====================================================

Write-Host "Collecting Enterprise Applications and SSO Applications..." -ForegroundColor Cyan

$Apps = Get-MgServicePrincipal -All -Property DisplayName,AppId,Id,AccountEnabled,ServicePrincipalType,PreferredSingleSignOnMode,PublisherName,AppOwnerOrganizationId,SignInAudience,Tags,Homepage,LoginUrl,LogoutUrl |
Select-Object DisplayName,
              AppId,
              @{Name="ObjectId";Expression={$_.Id}},
              AccountEnabled,
              ServicePrincipalType,
              PreferredSingleSignOnMode,
              PublisherName,
              AppOwnerOrganizationId,
              SignInAudience,
              @{Name="Tags";Expression={$_.Tags -join "; "}},
              Homepage,
              LoginUrl,
              LogoutUrl

$Apps = $Apps | Sort-Object DisplayName
$SSOApps = $Apps | Where-Object { $_.PreferredSingleSignOnMode -ne $null -and $_.PreferredSingleSignOnMode -ne "" } | Sort-Object DisplayName

$EnterpriseAppsPath = Export-ReportCsv -Data $Apps -FileName "EnterpriseApplications.csv"
$SSOAppsPath = Export-ReportCsv -Data $SSOApps -FileName "SSOApplications.csv"

Add-ExcelSheet -Data $Apps -WorksheetName "Enterprise Apps"
Add-ExcelSheet -Data $SSOApps -WorksheetName "SSO Apps"

# =====================================================
# Privileged & Excluded Accounts
# =====================================================

Write-Host "Collecting Privileged & Excluded Accounts..." -ForegroundColor Cyan

$KeywordCandidates = $Users | Where-Object {
    ($_.Description -match "break|emergency") -or
    ($_.DisplayName -match "break|emergency") -or
    ($_.UserPrincipalName -match "breakglass|break-glass|emergency|bg")
} | ForEach-Object {
    [PSCustomObject]@{
        Source            = "Name/UPN/Description Match"
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        Mail              = $_.Mail
        Description       = $_.Description
        AccountEnabled    = $_.AccountEnabled
    }
}

$CAExcludedCandidates = $CAExclusions | ForEach-Object {
    [PSCustomObject]@{
        Source            = "Conditional Access Exclusion"
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        Mail              = $_.Mail
        Description       = $_.Description
        AccountEnabled    = $_.AccountEnabled
    }
}

$GlobalAdminCandidates = $GlobalAdmins | ForEach-Object {
    [PSCustomObject]@{
        Source            = "Global Administrator"
        DisplayName       = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        Mail              = $_.Mail
        Description       = $_.Description
        AccountEnabled    = $_.AccountEnabled
    }
}

$PrivilegedExcludedAccounts = @($KeywordCandidates + $CAExcludedCandidates + $GlobalAdminCandidates) |
Where-Object { $_.UserPrincipalName -or $_.DisplayName } |
Sort-Object UserPrincipalName, Source -Unique

$PrivilegedExcludedPath = Export-ReportCsv -Data $PrivilegedExcludedAccounts -FileName "PrivilegedExcludedAccounts.csv"
Add-ExcelSheet -Data $PrivilegedExcludedAccounts -WorksheetName "Privileged & Excluded"

# =====================================================
# Report Index
# =====================================================

$ReportIndex = @(
    [PSCustomObject]@{ Report = "TenantSummary"; File = $TenantSummaryPath; Count = $TenantSummary.Count }
    [PSCustomObject]@{ Report = "ActiveUsers"; File = $ActiveUsersPath; Count = $ActiveUsers.Count }
    [PSCustomObject]@{ Report = "DisabledUsers"; File = $DisabledUsersPath; Count = $DisabledUsers.Count }
    [PSCustomObject]@{ Report = "LicensedUsers"; File = $LicensedUsersPath; Count = $LicensedUsers.Count }
    [PSCustomObject]@{ Report = "GuestUsers"; File = $GuestUsersPath; Count = $GuestUsers.Count }
    [PSCustomObject]@{ Report = "DirectoryRoleAssignments"; File = $DirectoryRolesPath; Count = $RoleAssignments.Count }
    [PSCustomObject]@{ Report = "GlobalAdmins"; File = $GlobalAdminsPath; Count = $GlobalAdmins.Count }
    [PSCustomObject]@{ Report = "ConditionalAccessPolicies"; File = $CAPoliciesPath; Count = $CAPolicies.Count }
    [PSCustomObject]@{ Report = "ConditionalAccessExclusions"; File = $CAExclusionsPath; Count = $CAExclusions.Count }
    [PSCustomObject]@{ Report = "NamedLocations"; File = $NamedLocationsPath; Count = $NamedLocations.Count }
    [PSCustomObject]@{ Report = "EnterpriseApplications"; File = $EnterpriseAppsPath; Count = $Apps.Count }
    [PSCustomObject]@{ Report = "SSOApplications"; File = $SSOAppsPath; Count = $SSOApps.Count }
    [PSCustomObject]@{ Report = "PrivilegedExcludedAccounts"; File = $PrivilegedExcludedPath; Count = $PrivilegedExcludedAccounts.Count }
)

$ReportIndexPath = Export-ReportCsv -Data $ReportIndex -FileName "ReportIndex.csv"
Add-ExcelSheet -Data $ReportIndex -WorksheetName "Report Index"

# Apply color-coded tabs after workbook generation.
Set-ExcelTabColors -Path $ExcelFile

# =====================================================
# Complete
# =====================================================

Write-Host ""
Write-Host "Reports generated in: $ReportPath" -ForegroundColor Green
Write-Host ""
$ReportIndex | Format-Table Report, Count, File -AutoSize

if ($script:ExcelAvailable -eq $true) {
    Write-Host ""
    Write-Host "Excel workbook created:" -ForegroundColor Green
    Write-Host $ExcelFile
}
else {
    Write-Host ""
    Write-Host "Excel workbook was not created because ImportExcel was unavailable. CSV reports were created successfully." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Completed Successfully." -ForegroundColor Green
