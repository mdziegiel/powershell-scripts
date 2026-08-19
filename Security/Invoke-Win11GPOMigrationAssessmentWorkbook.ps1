<#
.SYNOPSIS
Creates the Windows 11 GPO Migration Assessment workbook.

.DESCRIPTION
This version does not depend on Get-GPInheritance for link discovery. It reads the AD gpLink attribute directly
from workstation OUs and the domain, then maps linked GPOs to Windows 10/11 devices by OU.

.REQUIREMENTS
- RSAT Active Directory module
- RSAT Group Policy module
- ImportExcel module
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = "C:\Reporting\GPO_Migration_Assessment",
    [string]$ReportName = "JDCU_Windows11_GPO_Migration_Assessment",
    [string[]]$SearchBase,
    [switch]$IncludeDisabledComputers,
    [switch]$IncludeServerOS
)

$ErrorActionPreference = "Stop"
$RunDate = Get-Date
$TimeStamp = $RunDate.ToString("yyyyMMdd-HHmmss")
$RunBy = "$env:USERDOMAIN\$env:USERNAME"
$LocalComputerName = $env:COMPUTERNAME
$OutputPath = Join-Path $OutputFolder "$ReportName`_$TimeStamp.xlsx"

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

function Import-RequiredModule {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$ValidationCommand,
        [string]$InstallHint
    )
    try {
        Import-Module $Name -ErrorAction Stop
        if (-not (Get-Command $ValidationCommand -ErrorAction SilentlyContinue)) {
            throw "Validation command '$ValidationCommand' was not found after importing '$Name'."
        }
    }
    catch {
        $msg = "Required module '$Name' could not be loaded. $($_.Exception.Message)"
        if ($InstallHint) { $msg += " Install hint: $InstallHint" }
        throw $msg
    }
}

Import-RequiredModule -Name ActiveDirectory -ValidationCommand Get-ADComputer -InstallHint "Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Import-RequiredModule -Name GroupPolicy -ValidationCommand Get-GPO -InstallHint "Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
Import-RequiredModule -Name ImportExcel -ValidationCommand Export-Excel -InstallHint "Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber"

$DomainInfo = Get-ADDomain
$DomainDN = $DomainInfo.DistinguishedName
$DomainDNS = $DomainInfo.DNSRoot

function Get-ParentContainerFromDN {
    param([string]$DistinguishedName)
    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return "Unknown" }
    $parts = $DistinguishedName -split ','
    if ($parts.Count -le 1) { return "Unknown" }
    return ($parts | Select-Object -Skip 1) -join ','
}

function Get-OSBucket {
    param([string]$OperatingSystem, [string]$OperatingSystemVersion)

    $osText = "$OperatingSystem $OperatingSystemVersion".ToLowerInvariant()
    if ($osText -match "windows 11") { return "Windows 11" }
    if ($osText -match "windows 10") { return "Windows 10" }
    if ($osText -match "windows 8|windows 7|windows xp") { return "Previous OS" }
    if ($osText -match "server") { return "Server OS" }
    if ($osText -match "windows") { return "Other Windows" }
    return "Unknown"
}

function Get-AncestorScopes {
    param([string]$ContainerDN)

    $scopes = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($ContainerDN) -or $ContainerDN -eq "Unknown") { return $scopes }

    $current = $ContainerDN
    while ($current) {
        $scopes.Add($current)
        if ($current -ieq $DomainDN) { break }
        $parts = $current -split ','
        if ($parts.Count -le 1) { break }
        $current = ($parts | Select-Object -Skip 1) -join ','
    }

    if (-not ($scopes -contains $DomainDN)) {
        $scopes.Add($DomainDN)
    }

    return $scopes
}

function Parse-GPLinkAttribute {
    param(
        [string]$ScopeDN,
        [string]$TargetContainerDN,
        [string]$GpLinkText,
        [int]$GpOptions
    )

    $rows = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($GpLinkText)) { return $rows }

    $matches = [regex]::Matches($GpLinkText, '\[LDAP://.*?\{(?<Guid>[0-9A-Fa-f\-]+)\}.*?;(?<Options>\d+)\]', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $matches) {
        $guidText = $m.Groups['Guid'].Value
        $optionValue = [int]$m.Groups['Options'].Value
        $linkEnabled = (($optionValue -band 1) -eq 0)
        $linkEnforced = (($optionValue -band 2) -eq 2)
        $scopeType = if ($ScopeDN -ieq $TargetContainerDN) { "Direct" } else { "Inherited" }
        $inheritanceBlockedAtScope = (($GpOptions -band 1) -eq 1)

        $rows.Add([pscustomobject]@{
            TargetContainerDN = $TargetContainerDN
            LinkScopeDN = $ScopeDN
            LinkScope = $scopeType
            GPOGuid = $guidText
            LinkEnabled = $linkEnabled
            LinkEnforced = $linkEnforced
            LinkOptions = $optionValue
            InheritanceBlockedAtScope = $inheritanceBlockedAtScope
        })
    }
    return $rows
}

function Get-MigrationDecision {
    param(
        [string]$GPOName,
        [int]$Windows10Count,
        [int]$Windows11Count,
        [int]$PreviousOSCount,
        [int]$TotalWindowsWorkstations
    )

    $n = $GPOName.ToLowerInvariant()

    if ($n -match "removal-pending|to be removed|delete only|^copy of|test|_testing|wpa3test") {
        return "Retire / Cleanup / Legacy"
    }
    if ($n -match "windows 10|w10 only|defer windows 11|ie domain|internet explorer|ie no proxy|compatibility|cached mode|workaround|needs review|trainingpc|server only|logmein") {
        return "Review / Rebuild / Manual Review"
    }
    if ($n -match "smb signing|ldap signing|channel bonding|aes|tls|cipher|laps|applocker|audit|ntlm|cve-|powershell 2.0|netbios|inactivity|local admins|rdp|certificate|root ca|sectigo|fortinet|umbrella root") {
        return "Bring Forward"
    }
    if ($n -match "chrome|edge|firefox|browser|extension|homepage|pdf|foxit|microsoft 365|outlook|autodiscover|signature|five9|adp|dna|printer|mapped|drive|share|folder|wifi|wireless|wired|802.1x|teap|umbrella|log360|netwrix|nautilus|verifast") {
        return "Bring Forward"
    }
    if ($Windows10Count -gt 0 -or $Windows11Count -gt 0) {
        return "Review / Rebuild / Manual Review"
    }
    if ($PreviousOSCount -gt 0 -and $TotalWindowsWorkstations -eq 0) {
        return "Retire / Cleanup / Legacy"
    }
    return "Review / Rebuild / Manual Review"
}

function Get-MigrationRationale {
    param([string]$Decision)
    switch ($Decision) {
        "Bring Forward" { return "Likely supports security, workstation baseline, certificates, browsers, apps, printers, drives, or network access. Validate before Windows 11 rollout." }
        "Review / Rebuild / Manual Review" { return "Needs review before migration. Could be Windows 10-specific, workaround-based, legacy app-related, or not clearly classifiable." }
        "Retire / Cleanup / Legacy" { return "Looks like a test, copy, removal-pending, legacy, or cleanup candidate. Confirm before retiring." }
        default { return "Manual review required." }
    }
}

Write-Host "Collecting Windows computer objects from Active Directory..." -ForegroundColor Cyan
$computerProps = @("OperatingSystem", "OperatingSystemVersion", "DistinguishedName", "Enabled", "LastLogonDate", "Description", "DNSHostName")
$rawComputers = New-Object System.Collections.Generic.List[object]

if ($SearchBase -and $SearchBase.Count -gt 0) {
    foreach ($base in $SearchBase) {
        Write-Host "SearchBase: $base" -ForegroundColor Gray
        Get-ADComputer -Filter * -SearchBase $base -Properties $computerProps | ForEach-Object { $rawComputers.Add($_) }
    }
}
else {
    Get-ADComputer -Filter * -Properties $computerProps | ForEach-Object { $rawComputers.Add($_) }
}

$WindowsDevices = foreach ($c in $rawComputers) {
    $osBucket = Get-OSBucket -OperatingSystem $c.OperatingSystem -OperatingSystemVersion $c.OperatingSystemVersion
    $isServerDevice = ($osBucket -eq "Server OS")
    $isWindowsDevice = ($osBucket -ne "Unknown")

    if (-not $IncludeDisabledComputers -and $c.Enabled -ne $true) { continue }
    if (-not $IncludeServerOS -and $isServerDevice) { continue }
    if (-not $isWindowsDevice) { continue }

    [pscustomobject]@{
        ComputerName = $c.Name
        DNSHostName = $c.DNSHostName
        Enabled = $c.Enabled
        OperatingSystem = $c.OperatingSystem
        OperatingSystemVersion = $c.OperatingSystemVersion
        OSBucket = $osBucket
        ParentContainerDN = Get-ParentContainerFromDN -DistinguishedName $c.DistinguishedName
        DistinguishedName = $c.DistinguishedName
        LastLogonDate = $c.LastLogonDate
        Description = $c.Description
    }
}
$WindowsDevices = @($WindowsDevices)
$UniqueContainers = @($WindowsDevices | Select-Object -ExpandProperty ParentContainerDN -Unique | Where-Object { $_ -and $_ -ne "Unknown" } | Sort-Object)

Write-Host "Found $($WindowsDevices.Count) Windows workstation objects." -ForegroundColor Green
Write-Host "Found $($UniqueContainers.Count) unique workstation containers/OUs." -ForegroundColor Green
Write-Host "Reading gpLink attributes directly from AD..." -ForegroundColor Cyan

$GpoGuidToName = @{}
$AllGpos = Get-GPO -All | Sort-Object DisplayName
foreach ($g in $AllGpos) {
    $GpoGuidToName[$g.Id.Guid.ToString().ToLowerInvariant()] = $g.DisplayName
}
$GpoLookupByName = @{}
foreach ($g in $AllGpos) { $GpoLookupByName[$g.DisplayName] = $g }

$ScopeCache = @{}
$OuLinkRows = New-Object System.Collections.Generic.List[object]
$RunLogRows = New-Object System.Collections.Generic.List[object]

foreach ($containerDN in $UniqueContainers) {
    Write-Host "Reading links for: $containerDN" -ForegroundColor Gray
    try {
        $ancestorScopes = Get-AncestorScopes -ContainerDN $containerDN
        foreach ($scopeDN in $ancestorScopes) {
            if (-not $ScopeCache.ContainsKey($scopeDN)) {
                try {
                    $scopeObj = Get-ADObject -Identity $scopeDN -Properties gpLink,gpOptions,objectClass,Name
                    $ScopeCache[$scopeDN] = $scopeObj
                }
                catch {
                    $ScopeCache[$scopeDN] = $null
                }
            }

            $scope = $ScopeCache[$scopeDN]
            if ($null -eq $scope) { continue }

            $parsedLinks = Parse-GPLinkAttribute -ScopeDN $scopeDN -TargetContainerDN $containerDN -GpLinkText $scope.gpLink -GpOptions ([int]$scope.gpOptions)
            foreach ($p in $parsedLinks) {
                $guidKey = $p.GPOGuid.ToLowerInvariant()
                $gpoName = if ($GpoGuidToName.ContainsKey($guidKey)) { $GpoGuidToName[$guidKey] } else { "Unknown GPO {$($p.GPOGuid)}" }
                $OuLinkRows.Add([pscustomobject]@{
                    TargetContainerDN = $p.TargetContainerDN
                    LinkScopeDN = $p.LinkScopeDN
                    LinkScope = $p.LinkScope
                    GPOName = $gpoName
                    GPOGuid = $p.GPOGuid
                    LinkEnabled = $p.LinkEnabled
                    LinkEnforced = $p.LinkEnforced
                    LinkOptions = $p.LinkOptions
                    InheritanceBlockedAtScope = $p.InheritanceBlockedAtScope
                })
            }
        }
        $RunLogRows.Add([pscustomobject]@{ Time = Get-Date; Item = $containerDN; Status = "Processed"; Notes = "" })
    }
    catch {
        $RunLogRows.Add([pscustomobject]@{ Time = Get-Date; Item = $containerDN; Status = "Error"; Notes = $_.Exception.Message })
    }
}

Write-Host "Enabled GPO links found: $(@($OuLinkRows | Where-Object LinkEnabled -eq $true).Count)" -ForegroundColor Green
Write-Host "Total GPO links found: $($OuLinkRows.Count)" -ForegroundColor Green

Write-Host "Building computer to GPO impact mapping..." -ForegroundColor Cyan
$AppliedGpoRows = New-Object System.Collections.Generic.List[object]
foreach ($device in $WindowsDevices) {
    $linksForContainer = @($OuLinkRows | Where-Object { $_.TargetContainerDN -eq $device.ParentContainerDN -and $_.LinkEnabled -eq $true })
    foreach ($link in $linksForContainer) {
        $AppliedGpoRows.Add([pscustomobject]@{
            ComputerName = $device.ComputerName
            DNSHostName = $device.DNSHostName
            OperatingSystem = $device.OperatingSystem
            OperatingSystemVersion = $device.OperatingSystemVersion
            OSBucket = $device.OSBucket
            ParentContainerDN = $device.ParentContainerDN
            GPOName = $link.GPOName
            GPOGuid = $link.GPOGuid
            LinkScope = $link.LinkScope
            LinkScopeDN = $link.LinkScopeDN
            LinkEnforced = $link.LinkEnforced
        })
    }
}

Write-Host "Building migration recommendations..." -ForegroundColor Cyan
$GpoImpactRows = New-Object System.Collections.Generic.List[object]
$grouped = $AppliedGpoRows | Group-Object GPOName | Sort-Object Name
foreach ($grp in $grouped) {
    $rows = @($grp.Group)
    $win10 = @($rows | Where-Object OSBucket -eq "Windows 10").Count
    $win11 = @($rows | Where-Object OSBucket -eq "Windows 11").Count
    $prev = @($rows | Where-Object OSBucket -eq "Previous OS").Count
    $other = @($rows | Where-Object OSBucket -eq "Other Windows").Count
    $total = $rows.Count
    $containerCount = @($rows | Select-Object -ExpandProperty ParentContainerDN -Unique).Count
    $gpo = $GpoLookupByName[$grp.Name]
    $decision = Get-MigrationDecision -GPOName $grp.Name -Windows10Count $win10 -Windows11Count $win11 -PreviousOSCount $prev -TotalWindowsWorkstations ($win10 + $win11 + $other)
    $rationale = Get-MigrationRationale -Decision $decision

    $GpoImpactRows.Add([pscustomobject]@{
        GPOName = $grp.Name
        GPOGuid = if ($gpo) { $gpo.Id } else { ($rows | Select-Object -First 1).GPOGuid }
        Owner = if ($gpo) { $gpo.Owner } else { "" }
        Created = if ($gpo) { $gpo.CreationTime } else { $null }
        Modified = if ($gpo) { $gpo.ModificationTime } else { $null }
        GPOStatus = if ($gpo) { $gpo.GpoStatus } else { "Unknown" }
        TotalAffectedWindowsDevices = $total
        AffectedWindows10Devices = $win10
        AffectedWindows11Devices = $win11
        AffectedPreviousOSDevices = $prev
        AffectedOtherWindowsDevices = $other
        AffectedContainerCount = $containerCount
        MigrationDecision = $decision
        MigrationRationale = $rationale
        ReviewOwner = ""
        FinalDecision = ""
        Notes = ""
    })
}

$ContainerSummaryRows = foreach ($containerDN in $UniqueContainers) {
    $devices = @($WindowsDevices | Where-Object { $_.ParentContainerDN -eq $containerDN })
    [pscustomobject]@{
        ContainerDN = $containerDN
        TotalWindowsDevices = $devices.Count
        Windows10Devices = @($devices | Where-Object OSBucket -eq "Windows 10").Count
        Windows11Devices = @($devices | Where-Object OSBucket -eq "Windows 11").Count
        PreviousOSDevices = @($devices | Where-Object OSBucket -eq "Previous OS").Count
        OtherWindowsDevices = @($devices | Where-Object OSBucket -eq "Other Windows").Count
        EnabledGpoLinksFound = @($OuLinkRows | Where-Object { $_.TargetContainerDN -eq $containerDN -and $_.LinkEnabled -eq $true }).Count
        TotalGpoLinksFound = @($OuLinkRows | Where-Object { $_.TargetContainerDN -eq $containerDN }).Count
    }
}

$BringForwardRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -eq "Bring Forward" })
$ReviewRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -eq "Review / Rebuild / Manual Review" })
$RetireRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -eq "Retire / Cleanup / Legacy" })
$IntuneCandidateRows = @($GpoImpactRows | Where-Object { $_.GPOName -match "Chrome|Edge|Firefox|Microsoft 365|Outlook|OneDrive|LAPS|BitLocker|Defender|Firewall|Update|Windows 11|Windows 10|Printer|Drive|Mapped|Wireless|Certificate|Browser|Extension" })

$ExecutiveSummaryRows = @(
    [pscustomobject]@{ Metric = "Report Run Date"; Value = $RunDate }
    [pscustomobject]@{ Metric = "Report Run By"; Value = $RunBy }
    [pscustomobject]@{ Metric = "Computer Name"; Value = $LocalComputerName }
    [pscustomobject]@{ Metric = "Domain"; Value = $DomainDNS }
    [pscustomobject]@{ Metric = "Windows Workstation Objects Reviewed"; Value = $WindowsDevices.Count }
    [pscustomobject]@{ Metric = "Windows 10 Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Windows 10").Count }
    [pscustomobject]@{ Metric = "Windows 11 Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Windows 11").Count }
    [pscustomobject]@{ Metric = "Previous OS Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Previous OS").Count }
    [pscustomobject]@{ Metric = "Workstation Containers Reviewed"; Value = $UniqueContainers.Count }
    [pscustomobject]@{ Metric = "Enabled GPO Links Found"; Value = @($OuLinkRows | Where-Object LinkEnabled -eq $true).Count }
    [pscustomobject]@{ Metric = "GPOs Affecting Windows Devices"; Value = $GpoImpactRows.Count }
    [pscustomobject]@{ Metric = "Bring Forward"; Value = $BringForwardRows.Count }
    [pscustomobject]@{ Metric = "Review / Rebuild / Manual Review"; Value = $ReviewRows.Count }
    [pscustomobject]@{ Metric = "Retire / Cleanup / Legacy"; Value = $RetireRows.Count }
    [pscustomobject]@{ Metric = "Potential Intune Candidates"; Value = $IntuneCandidateRows.Count }
)

$ReportIndexRows = @(
    [pscustomobject]@{ Tab = "Report Index"; Purpose = "Workbook map and tab descriptions"; RowCount = 15 }
    [pscustomobject]@{ Tab = "Executive Summary"; Purpose = "High-level Windows 11 migration counts"; RowCount = $ExecutiveSummaryRows.Count }
    [pscustomobject]@{ Tab = "Windows Devices"; Purpose = "All Windows workstation computer objects reviewed"; RowCount = $WindowsDevices.Count }
    [pscustomobject]@{ Tab = "Windows 10 Devices"; Purpose = "Windows 10 devices found in AD"; RowCount = @($WindowsDevices | Where-Object OSBucket -eq "Windows 10").Count }
    [pscustomobject]@{ Tab = "Windows 11 Devices"; Purpose = "Windows 11 devices found in AD"; RowCount = @($WindowsDevices | Where-Object OSBucket -eq "Windows 11").Count }
    [pscustomobject]@{ Tab = "Previous OS Devices"; Purpose = "Windows XP/7/8 or other previous OS devices found in AD"; RowCount = @($WindowsDevices | Where-Object OSBucket -eq "Previous OS").Count }
    [pscustomobject]@{ Tab = "Workstation Containers"; Purpose = "Container/OU-level Windows device counts and GPO link summary"; RowCount = $ContainerSummaryRows.Count }
    [pscustomobject]@{ Tab = "Container GPO Links"; Purpose = "Direct and inherited GPO links parsed from AD gpLink attributes"; RowCount = $OuLinkRows.Count }
    [pscustomobject]@{ Tab = "Applied GPOs"; Purpose = "Computer-to-GPO impact map"; RowCount = $AppliedGpoRows.Count }
    [pscustomobject]@{ Tab = "GPO Impact Summary"; Purpose = "GPO-level affected device counts and migration recommendation"; RowCount = $GpoImpactRows.Count }
    [pscustomobject]@{ Tab = "Bring Forward"; Purpose = "Policies expected to migrate to Windows 11 after validation"; RowCount = $BringForwardRows.Count }
    [pscustomobject]@{ Tab = "Review Rebuild"; Purpose = "Policies needing review, rebuild, or manual decision"; RowCount = $ReviewRows.Count }
    [pscustomobject]@{ Tab = "Retire Cleanup"; Purpose = "Policies that look temporary, legacy, test, or retirement candidates"; RowCount = $RetireRows.Count }
    [pscustomobject]@{ Tab = "Intune Candidates"; Purpose = "Policies that may be candidates to move to Intune later"; RowCount = $IntuneCandidateRows.Count }
    [pscustomobject]@{ Tab = "Run Log"; Purpose = "Processing log"; RowCount = $RunLogRows.Count }
)

if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$excelParams = @{ Path = $OutputPath; AutoSize = $true; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true }

$ReportIndexRows | Export-Excel @excelParams -WorksheetName "Report Index" -TableName "tblReportIndex" -ClearSheet
$ExecutiveSummaryRows | Export-Excel @excelParams -WorksheetName "Executive Summary" -TableName "tblExecutiveSummary" -ClearSheet
$WindowsDevices | Export-Excel @excelParams -WorksheetName "Windows Devices" -TableName "tblWindowsDevices" -ClearSheet
$WindowsDevices | Where-Object OSBucket -eq "Windows 10" | Export-Excel @excelParams -WorksheetName "Windows 10 Devices" -TableName "tblWindows10Devices" -ClearSheet
$WindowsDevices | Where-Object OSBucket -eq "Windows 11" | Export-Excel @excelParams -WorksheetName "Windows 11 Devices" -TableName "tblWindows11Devices" -ClearSheet
$WindowsDevices | Where-Object OSBucket -eq "Previous OS" | Export-Excel @excelParams -WorksheetName "Previous OS Devices" -TableName "tblPreviousOSDevices" -ClearSheet
$ContainerSummaryRows | Export-Excel @excelParams -WorksheetName "Workstation Containers" -TableName "tblWorkstationContainers" -ClearSheet
$OuLinkRows | Export-Excel @excelParams -WorksheetName "Container GPO Links" -TableName "tblContainerGpoLinks" -ClearSheet
$AppliedGpoRows | Export-Excel @excelParams -WorksheetName "Applied GPOs" -TableName "tblAppliedGPOs" -ClearSheet
$GpoImpactRows | Export-Excel @excelParams -WorksheetName "GPO Impact Summary" -TableName "tblGPOImpactSummary" -ClearSheet
$BringForwardRows | Export-Excel @excelParams -WorksheetName "Bring Forward" -TableName "tblBringForward" -ClearSheet
$ReviewRows | Export-Excel @excelParams -WorksheetName "Review Rebuild" -TableName "tblReviewRebuild" -ClearSheet
$RetireRows | Export-Excel @excelParams -WorksheetName "Retire Cleanup" -TableName "tblRetireCleanup" -ClearSheet
$IntuneCandidateRows | Export-Excel @excelParams -WorksheetName "Intune Candidates" -TableName "tblIntuneCandidates" -ClearSheet
$RunLogRows | Export-Excel @excelParams -WorksheetName "Run Log" -TableName "tblRunLog" -ClearSheet

$pkg = Open-ExcelPackage -Path $OutputPath
foreach ($ws in $pkg.Workbook.Worksheets) {
    $ws.View.ShowGridLines = $false
    if ($ws.Dimension) {
        $lastCol = $ws.Dimension.End.Column
        $lastRow = $ws.Dimension.End.Row
        $headerRange = $ws.Cells[1,1,1,$lastCol]
        $headerRange.Style.Font.Bold = $true
        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $headerRange.Style.Fill.PatternType = 'Solid'
        $headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31,78,121))
        $ws.Cells[1,1,$lastRow,$lastCol].Style.VerticalAlignment = 'Top'
        $ws.Cells[1,1,$lastRow,$lastCol].Style.WrapText = $false
        $ws.Row(1).Height = 22
    }
}
Close-ExcelPackage $pkg

Write-Host "Windows 11 GPO Migration Assessment workbook created:" -ForegroundColor Green
Write-Host $OutputPath -ForegroundColor Green
