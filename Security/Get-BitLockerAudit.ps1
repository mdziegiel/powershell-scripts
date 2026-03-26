<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get BitLocker Audit
SYNOPSIS    : Audits BitLocker encryption status across domain computers
DESCRIPTION : Queries Active Directory computers remotely to collect
              BitLocker encryption status for all drives. Identifies
              unencrypted or partially encrypted systems for remediation
              and exports results for reporting
==============================================================================
#>

[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [string]$OUFilter,
    [switch]$UnprotectedOnly
)

$ErrorActionPreference = "SilentlyContinue"
$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$ReportPath = Join-Path $OutputPath "BitLockerAudit_$Timestamp.csv"
$Results = [System.Collections.Generic.List[PSObject]]::new()

# Get computers from AD
Write-Host "Querying Active Directory for computer accounts..." -ForegroundColor Cyan
$ADParams = @{ Filter = "*"; Properties = "Name", "OperatingSystem", "LastLogonDate" }
if ($OUFilter) { $ADParams["SearchBase"] = $OUFilter }
$Computers = Get-ADComputer @ADParams | Where-Object { $_.OperatingSystem -like "*Windows*" }
Write-Host "Found $($Computers.Count) computers. Checking BitLocker status..." -ForegroundColor Cyan

$i = 0
foreach ($Computer in $Computers) {
    $i++
    Write-Progress -Activity "Checking BitLocker Status" -Status "$($Computer.Name) ($i of $($Computers.Count))" -PercentComplete (($i / $Computers.Count) * 100)

    if (-not (Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet)) {
        $Results.Add([PSCustomObject]@{
            ComputerName      = $Computer.Name
            OS                = $Computer.OperatingSystem
            LastLogonDate     = $Computer.LastLogonDate
            DriveLetter       = "N/A"
            VolumeStatus      = "UNREACHABLE"
            ProtectionStatus  = "UNREACHABLE"
            EncryptionMethod  = "N/A"
            KeyProtectors     = "N/A"
            AuditTime         = Get-Date -Format "yyyy-MM-dd HH:mm"
        })
        continue
    }

    try {
        $BLStatus = Invoke-Command -ComputerName $Computer.Name -ScriptBlock {
            Get-BitLockerVolume | Select-Object MountPoint, VolumeStatus, ProtectionStatus, EncryptionMethod,
                @{N="KeyProtectors"; E={ ($_.KeyProtector | Select-Object -ExpandProperty KeyProtectorType) -join "; " }}
        } -ErrorAction Stop

        foreach ($Vol in $BLStatus) {
            $Results.Add([PSCustomObject]@{
                ComputerName      = $Computer.Name
                OS                = $Computer.OperatingSystem
                LastLogonDate     = $Computer.LastLogonDate
                DriveLetter       = $Vol.MountPoint
                VolumeStatus      = $Vol.VolumeStatus
                ProtectionStatus  = $Vol.ProtectionStatus
                EncryptionMethod  = $Vol.EncryptionMethod
                KeyProtectors     = $Vol.KeyProtectors
                AuditTime         = Get-Date -Format "yyyy-MM-dd HH:mm"
            })
        }
    }
    catch {
        $Results.Add([PSCustomObject]@{
            ComputerName      = $Computer.Name
            OS                = $Computer.OperatingSystem
            LastLogonDate     = $Computer.LastLogonDate
            DriveLetter       = "N/A"
            VolumeStatus      = "ERROR: $($_.Exception.Message)"
            ProtectionStatus  = "N/A"
            EncryptionMethod  = "N/A"
            KeyProtectors     = "N/A"
            AuditTime         = Get-Date -Format "yyyy-MM-dd HH:mm"
        })
    }
}

Write-Progress -Activity "Checking BitLocker Status" -Completed

$Output = if ($UnprotectedOnly) {
    $Results | Where-Object { $_.ProtectionStatus -ne "On" }
} else { $Results }

$Output | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

$Unprotected = ($Results | Where-Object { $_.ProtectionStatus -eq "Off" }).Count
Write-Host "`nAudit complete. $($Results.Count) volumes checked." -ForegroundColor Green
Write-Host "Unprotected volumes: $Unprotected" -ForegroundColor $(if ($Unprotected -gt 0) { "Red" } else { "Green" })
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow
