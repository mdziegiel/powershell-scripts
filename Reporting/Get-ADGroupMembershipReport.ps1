#Requires -Version 5.1
<#
.SYNOPSIS
    Exports Active Directory group memberships for all users or a specific group.
.DESCRIPTION
    Generates a flat CSV of group memberships in hk.lan. Can run against all groups,
    a specific group, or export a per-user view of all group memberships.
    Useful for access reviews and offboarding audits.
.PARAMETER GroupName
    Optional. Specific AD group name to audit. If omitted, audits all groups.
.PARAMETER OutputPath
    Path for the CSV output file. Defaults to script directory.
.PARAMETER ByUser
    Switch to output one row per user showing all their group memberships.
.EXAMPLE
    .\Get-ADGroupMembershipReport.ps1
    .\Get-ADGroupMembershipReport.ps1 -GroupName "VPN-Users"
    .\Get-ADGroupMembershipReport.ps1 -ByUser -OutputPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [string]$GroupName,
    [string]$OutputPath = $PSScriptRoot,
    [switch]$ByUser
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$Results = [System.Collections.Generic.List[PSObject]]::new()

if ($ByUser) {
    # Per-user view: one row per user with all groups listed
    $ReportPath = Join-Path $OutputPath "ADGroupMembership_ByUser_$Timestamp.csv"
    Write-Host "Querying all users and their group memberships..." -ForegroundColor Cyan

    $Users = Get-ADUser -Filter { Enabled -eq $true } -Properties MemberOf, DisplayName, Department, Title

    foreach ($User in $Users) {
        $Groups = $User.MemberOf | ForEach-Object {
            (Get-ADGroup $_).Name
        }
        $Results.Add([PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            UserPrincipalName = $User.UserPrincipalName
            Department        = $User.Department
            Title             = $User.Title
            GroupCount        = $Groups.Count
            Groups            = $Groups -join "; "
        })
    }
}
else {
    # Per-group view: one row per group member
    $ReportPath = Join-Path $OutputPath "ADGroupMembership_ByGroup_$Timestamp.csv"

    $Groups = if ($GroupName) {
        Get-ADGroup -Filter { Name -eq $GroupName } -Properties Description
    } else {
        Get-ADGroup -Filter "*" -Properties Description
    }

    Write-Host "Auditing $($Groups.Count) group(s)..." -ForegroundColor Cyan
    $i = 0

    foreach ($Group in $Groups) {
        $i++
        Write-Progress -Activity "Auditing Group Memberships" -Status "$($Group.Name) ($i of $($Groups.Count))" -PercentComplete (($i / $Groups.Count) * 100)

        try {
            $Members = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction Stop
            foreach ($Member in $Members) {
                $MemberDetails = Get-ADObject -Identity $Member.DistinguishedName -Properties DisplayName, Department, Title, Enabled
                $Results.Add([PSCustomObject]@{
                    GroupName        = $Group.Name
                    GroupDescription = $Group.Description
                    GroupCategory    = $Group.GroupCategory
                    MemberName       = $Member.Name
                    MemberSAM        = $Member.SamAccountName
                    MemberType       = $Member.objectClass
                    Department       = $MemberDetails.Department
                    Title            = $MemberDetails.Title
                    Enabled          = $MemberDetails.Enabled
                })
            }
        }
        catch {
            $Results.Add([PSCustomObject]@{
                GroupName        = $Group.Name
                GroupDescription = $Group.Description
                GroupCategory    = $Group.GroupCategory
                MemberName       = "ERROR: $($_.Exception.Message)"
                MemberSAM        = ""
                MemberType       = ""
                Department       = ""
                Title            = ""
                Enabled          = ""
            })
        }
    }
    Write-Progress -Activity "Auditing Group Memberships" -Completed
}

$Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nExported $($Results.Count) records to: $ReportPath" -ForegroundColor Green
