#Requires -Version 5.1
<#
.SYNOPSIS
    Audits local Administrators group membership across all domain-joined computers.
.DESCRIPTION
    Pulls all computer accounts from hk.lan Active Directory, connects to each machine
    remotely, and enumerates the local Administrators group. Outputs results to CSV.
.PARAMETER OutputPath
    Path for the CSV output file. Defaults to script directory.
.PARAMETER OUFilter
    Optional OU to scope the computer search (DistinguishedName format).
.EXAMPLE
    .\Get-LocalAdminAudit.ps1
    .\Get-LocalAdminAudit.ps1 -OutputPath "C:\Reports" -OUFilter "OU=Workstations,DC=hk,DC=lan"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [string]$OUFilter
)

$ErrorActionPreference = "SilentlyContinue"
$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "LocalAdminAudit_$Timestamp.csv"
$Results = [System.Collections.Generic.List[PSObject]]::new()

# Get computers from AD
Write-Host "Querying Active Directory for computer accounts..." -ForegroundColor Cyan
$ADParams = @{ Filter = "*"; Properties = "Name", "OperatingSystem", "LastLogonDate" }
if ($OUFilter) { $ADParams["SearchBase"] = $OUFilter }
$Computers = Get-ADComputer @ADParams | Where-Object { $_.OperatingSystem -like "*Windows*" }
Write-Host "Found $($Computers.Count) computers. Starting audit..." -ForegroundColor Cyan

$i = 0
foreach ($Computer in $Computers) {
    $i++
    Write-Progress -Activity "Auditing Local Admins" -Status "$($Computer.Name) ($i of $($Computers.Count))" -PercentComplete (($i / $Computers.Count) * 100)

    # Test connectivity
    if (-not (Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet)) {
        $Results.Add([PSCustomObject]@{
            ComputerName  = $Computer.Name
            OS            = $Computer.OperatingSystem
            LastLogonDate = $Computer.LastLogonDate
            MemberName    = "UNREACHABLE"
            MemberType    = "N/A"
            AuditTime     = Get-Date -Format "yyyy-MM-dd HH:mm"
        })
        continue
    }

    try {
        $Members = Invoke-Command -ComputerName $Computer.Name -ScriptBlock {
            $group = [ADSI]"WinNT://./Administrators,group"
            $group.Members() | ForEach-Object {
                $path = $_.GetType().InvokeMember("ADsPath", "GetProperty", $null, $_, $null)
                $name = $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null)
                [PSCustomObject]@{
                    Name = $name
                    Path = $path
                }
            }
        } -ErrorAction Stop

        foreach ($Member in $Members) {
            $type = if ($Member.Path -like "*WinNT*") { "Local" } else { "Domain" }
            $Results.Add([PSCustomObject]@{
                ComputerName  = $Computer.Name
                OS            = $Computer.OperatingSystem
                LastLogonDate = $Computer.LastLogonDate
                MemberName    = $Member.Name
                MemberType    = $type
                AuditTime     = Get-Date -Format "yyyy-MM-dd HH:mm"
            })
        }
    }
    catch {
        $Results.Add([PSCustomObject]@{
            ComputerName  = $Computer.Name
            OS            = $Computer.OperatingSystem
            LastLogonDate = $Computer.LastLogonDate
            MemberName    = "ACCESS DENIED / ERROR: $($_.Exception.Message)"
            MemberType    = "N/A"
            AuditTime     = Get-Date -Format "yyyy-MM-dd HH:mm"
        })
    }
}

Write-Progress -Activity "Auditing Local Admins" -Completed
$Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nAudit complete. $($Results.Count) records written to:" -ForegroundColor Green
Write-Host $ReportPath -ForegroundColor Yellow
