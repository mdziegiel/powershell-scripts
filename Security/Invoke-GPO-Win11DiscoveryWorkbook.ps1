<#
.SYNOPSIS
Creates the second workbook: Windows 11 GPO Migration Assessment.

.DESCRIPTION
This report maps Windows workstation computer objects to their OU, reads GPO links/inheritance for those OUs,
and builds a decision workbook showing which GPOs should be brought forward, reviewed/rebuilt, retired,
or considered for Intune migration during the Windows 11 project.

.REQUIREMENTS
- Run from a domain-joined admin workstation/server
- RSAT Active Directory PowerShell tools
- RSAT Group Policy Management tools
- ImportExcel PowerShell module

Install prerequisites if needed:
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0
Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber

.NOTES
This workbook answers: What GPOs are actively linked to Windows workstations and what should migrate to Windows 11?
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
$ComputerName = $env:COMPUTERNAME
$OutputPath = Join-Path $OutputFolder "$ReportName`_$TimeStamp.xlsx"

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

function Import-RequiredModule {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [string]$ValidationCommand,
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

try {
    $DomainInfo = Get-ADDomain -ErrorAction Stop
}
catch {
    $DomainInfo = $null
}

function Get-ParentOUFromDN {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return "Unknown" }
    $parts = $DistinguishedName -split ','
    if ($parts.Count -le 1) { return "Unknown" }
    return ($parts | Select-Object -Skip 1) -join ','
}

function Get-OSBucket {
    param([string]$OperatingSystem, [string]$OperatingSystemVersion)

    $os = "$OperatingSystem $OperatingSystemVersion".ToLowerInvariant()
    if ($os -match "windows 11") { return "Windows 11" }
    if ($os -match "windows 10") { return "Windows 10" }
    if ($os -match "windows 8") { return "Previous OS" }
    if ($os -match "windows 7") { return "Previous OS" }
    if ($os -match "windows xp") { return "Previous OS" }
    if ($os -match "server") { return "Server OS" }
    if ($os -match "windows") { return "Other Windows" }
    return "Unknown"
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
        return "Retire / cleanup candidate"
    }

    if ($n -match "windows 10|w10 only|defer windows 11|ie domain|internet explorer|ie no proxy|compatibility|cached mode|workaround|needs review|trainingpc|server only|logmein") {
        return "Review / rebuild for Windows 11"
    }

    if ($n -match "smb signing|ldap signing|channel bonding|aes|tls|cipher|laps|applocker|audit|ntlm|cve-|powershell 2.0|netbios|inactivity|local admins|rdp|certificate|root ca|sectigo|fortinet|umbrella root") {
        return "Bring forward"
    }

    if ($n -match "chrome|edge|firefox|browser|extension|homepage|pdf|foxit|microsoft 365|outlook|autodiscover|signature|five9|adp|dna|printer|mapped|drive|share|folder|wifi|wireless|wired|802.1x|teap|umbrella|log360|netwrix|nautilus|verifast") {
        return "Bring forward"
    }

    if ($Windows10Count -gt 0 -or $Windows11Count -gt 0) {
        return "Review for Windows 11"
    }

    if ($PreviousOSCount -gt 0 -and $TotalWindowsWorkstations -eq 0) {
        return "Legacy only / review for retirement"
    }

    return "Manual review"
}

function Get-MigrationRationale {
    param(
        [string]$GPOName,
        [string]$Decision,
        [int]$Windows10Count,
        [int]$Windows11Count,
        [int]$PreviousOSCount,
        [int]$TotalComputerCount
    )

    switch ($Decision) {
        "Bring forward" {
            return "Policy appears to support security, workstation baseline, certificates, browsers, applications, printers, drives, or network access. Validate settings before production Windows 11 rollout."
        }
        "Review / rebuild for Windows 11" {
            return "Policy name suggests Windows 10, IE, compatibility, workaround, server-only, or explicit review scope. Validate whether it is still required or should be rebuilt for Windows 11."
        }
        "Retire / cleanup candidate" {
            return "Policy name suggests test, copy, removal pending, delete-only, or temporary state. Confirm no current business owner before retirement."
        }
        "Legacy only / review for retirement" {
            return "Policy appears linked only to previous OS devices based on discovered AD computer OS data. Confirm before retirement."
        }
        default {
            return "Policy affects Windows devices or could not be confidently classified. Manual review required."
        }
    }
}

function Normalize-GpoLinkRows {
    param(
        [object[]]$Links,
        [string]$OU,
        [string]$Scope
    )

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($link in @($Links)) {
        if ($null -eq $link) { continue }

        $displayName = $link.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $link.Target }
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $link.Name }

        $enabled = $link.Enabled
        $enforced = $link.Enforced
        if ($null -eq $enforced) { $enforced = $link.NoOverride }

        $gpoId = $link.GpoId
        if ($null -eq $gpoId) { $gpoId = $link.Id }

        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $rows.Add([pscustomobject]@{
                OU = $OU
                LinkScope = $Scope
                GPOName = $displayName
                GPOGuid = $gpoId
                LinkEnabled = $enabled
                LinkEnforced = $enforced
            })
        }
    }
    return $rows
}

Write-Host "Collecting Windows computer objects from Active Directory..." -ForegroundColor Cyan

$computerProps = @("OperatingSystem", "OperatingSystemVersion", "DistinguishedName", "Enabled", "LastLogonDate", "Description", "DNSHostName")
$computerRows = New-Object System.Collections.Generic.List[object]

if ($SearchBase -and $SearchBase.Count -gt 0) {
    foreach ($base in $SearchBase) {
        Write-Host "SearchBase: $base" -ForegroundColor Gray
        $items = Get-ADComputer -Filter * -SearchBase $base -Properties $computerProps
        foreach ($c in $items) { $computerRows.Add($c) }
    }
}
else {
    $items = Get-ADComputer -Filter * -Properties $computerProps
    foreach ($c in $items) { $computerRows.Add($c) }
}

$WindowsDevices = foreach ($c in $computerRows) {
    $osBucket = Get-OSBucket -OperatingSystem $c.OperatingSystem -OperatingSystemVersion $c.OperatingSystemVersion
    $isServer = ($osBucket -eq "Server OS")
    $isWindowsDevice = ($osBucket -ne "Unknown")

    if (-not $IncludeDisabledComputers -and $c.Enabled -ne $true) { continue }
    if (-not $IncludeServerOS -and $isServer) { continue }
    if (-not $isWindowsDevice) { continue }

    [pscustomobject]@{
        ComputerName = $c.Name
        DNSHostName = $c.DNSHostName
        Enabled = $c.Enabled
        OperatingSystem = $c.OperatingSystem
        OperatingSystemVersion = $c.OperatingSystemVersion
        OSBucket = $osBucket
        ParentOU = Get-ParentOUFromDN -DistinguishedName $c.DistinguishedName
        DistinguishedName = $c.DistinguishedName
        LastLogonDate = $c.LastLogonDate
        Description = $c.Description
    }
}

$WindowsDevices = @($WindowsDevices)
$UniqueOUs = @($WindowsDevices | Select-Object -ExpandProperty ParentOU -Unique | Where-Object { $_ -and $_ -ne "Unknown" } | Sort-Object)

Write-Host "Found $($WindowsDevices.Count) Windows workstation objects." -ForegroundColor Green
Write-Host "Found $($UniqueOUs.Count) unique workstation OUs." -ForegroundColor Green
Write-Host "Collecting GPO inheritance for workstation OUs..." -ForegroundColor Cyan

$OuLinkRows = New-Object System.Collections.Generic.List[object]
$OuSummaryRows = New-Object System.Collections.Generic.List[object]
$RunLogRows = New-Object System.Collections.Generic.List[object]

foreach ($ou in $UniqueOUs) {
    Write-Host "Reading GPO inheritance: $ou" -ForegroundColor Gray
    try {
        $inheritance = Get-GPInheritance -Target $ou
        $direct = Normalize-GpoLinkRows -Links $inheritance.GpoLinks -OU $ou -Scope "Direct"
        $inherited = Normalize-GpoLinkRows -Links $inheritance.InheritedGpoLinks -OU $ou -Scope "Inherited"

        foreach ($r in $direct) { $OuLinkRows.Add($r) }
        foreach ($r in $inherited) { $OuLinkRows.Add($r) }

        $devicesInOu = @($WindowsDevices | Where-Object { $_.ParentOU -eq $ou })
        $OuSummaryRows.Add([pscustomobject]@{
            OU = $ou
            TotalWindowsDevices = $devicesInOu.Count
            Windows10Devices = @($devicesInOu | Where-Object OSBucket -eq "Windows 10").Count
            Windows11Devices = @($devicesInOu | Where-Object OSBucket -eq "Windows 11").Count
            PreviousOSDevices = @($devicesInOu | Where-Object OSBucket -eq "Previous OS").Count
            OtherWindowsDevices = @($devicesInOu | Where-Object OSBucket -eq "Other Windows").Count
            GPOInheritanceBlocked = $inheritance.GpoInheritanceBlocked
            DirectGpoLinks = @($direct).Count
            InheritedGpoLinks = @($inherited).Count
        })

        $RunLogRows.Add([pscustomobject]@{ Time = Get-Date; Item = $ou; Status = "Processed"; Notes = "" })
    }
    catch {
        $RunLogRows.Add([pscustomobject]@{ Time = Get-Date; Item = $ou; Status = "Error"; Notes = $_.Exception.Message })
    }
}

Write-Host "Building computer to GPO impact mapping..." -ForegroundColor Cyan

$AppliedGpoRows = New-Object System.Collections.Generic.List[object]
foreach ($device in $WindowsDevices) {
    $linksForOu = @($OuLinkRows | Where-Object { $_.OU -eq $device.ParentOU })
    foreach ($link in $linksForOu) {
        $enabledText = "$($link.LinkEnabled)"
        if ($enabledText -match "False|No|Disabled") { continue }
        $AppliedGpoRows.Add([pscustomobject]@{
            ComputerName = $device.ComputerName
            DNSHostName = $device.DNSHostName
            OperatingSystem = $device.OperatingSystem
            OperatingSystemVersion = $device.OperatingSystemVersion
            OSBucket = $device.OSBucket
            ParentOU = $device.ParentOU
            GPOName = $link.GPOName
            GPOGuid = $link.GPOGuid
            LinkScope = $link.LinkScope
            LinkEnforced = $link.LinkEnforced
        })
    }
}

$AllGpos = Get-GPO -All | Sort-Object DisplayName
$GpoLookup = @{}
foreach ($g in $AllGpos) {
    $GpoLookup[$g.DisplayName] = $g
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
    $ouCount = @($rows | Select-Object -ExpandProperty ParentOU -Unique).Count
    $gpo = $GpoLookup[$grp.Name]

    $decision = Get-MigrationDecision -GPOName $grp.Name -Windows10Count $win10 -Windows11Count $win11 -PreviousOSCount $prev -TotalWindowsWorkstations ($win10 + $win11 + $other)
    $rationale = Get-MigrationRationale -GPOName $grp.Name -Decision $decision -Windows10Count $win10 -Windows11Count $win11 -PreviousOSCount $prev -TotalComputerCount $total

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
        AffectedOUCount = $ouCount
        MigrationDecision = $decision
        MigrationRationale = $rationale
        ReviewOwner = ""
        FinalDecision = ""
        Notes = ""
    })
}

$BringForwardRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -eq "Bring forward" })
$ReviewRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -match "Review|Manual" })
$RetireRows = @($GpoImpactRows | Where-Object { $_.MigrationDecision -match "Retire|Legacy" })
$IntuneCandidateRows = @($GpoImpactRows | Where-Object { $_.GPOName -match "Chrome|Edge|Firefox|Microsoft 365|Outlook|OneDrive|LAPS|BitLocker|Defender|Firewall|Update|Windows 11|Windows 10|Printer|Drive|Mapped|Wireless|Certificate|Browser|Extension" })

$ExecutiveSummaryRows = @(
    [pscustomobject]@{ Metric = "Report Run Date"; Value = $RunDate }
    [pscustomobject]@{ Metric = "Report Run By"; Value = $RunBy }
    [pscustomobject]@{ Metric = "Computer Name"; Value = $ComputerName }
    [pscustomobject]@{ Metric = "Domain"; Value = if ($DomainInfo) { $DomainInfo.DNSRoot } else { "Not available" } }
    [pscustomobject]@{ Metric = "Windows Workstation Objects Reviewed"; Value = $WindowsDevices.Count }
    [pscustomobject]@{ Metric = "Windows 10 Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Windows 10").Count }
    [pscustomobject]@{ Metric = "Windows 11 Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Windows 11").Count }
    [pscustomobject]@{ Metric = "Previous OS Devices"; Value = @($WindowsDevices | Where-Object OSBucket -eq "Previous OS").Count }
    [pscustomobject]@{ Metric = "Workstation OUs Reviewed"; Value = $UniqueOUs.Count }
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
    [pscustomobject]@{ Tab = "Workstation OUs"; Purpose = "OU-level Windows device counts and GPO inheritance summary"; RowCount = $OuSummaryRows.Count }
    [pscustomobject]@{ Tab = "OU GPO Links"; Purpose = "Direct and inherited GPO links for workstation OUs"; RowCount = $OuLinkRows.Count }
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
$OuSummaryRows | Export-Excel @excelParams -WorksheetName "Workstation OUs" -TableName "tblWorkstationOUs" -ClearSheet
$OuLinkRows | Export-Excel @excelParams -WorksheetName "OU GPO Links" -TableName "tblOUGPOLinks" -ClearSheet
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
