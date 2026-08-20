<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Optimize-WindowsDisk
SYNOPSIS    : Audits Windows storage and optionally runs supported disk
              optimization and cleanup tasks.
DESCRIPTION : Audits Windows storage and optionally runs supported disk
              optimization and cleanup tasks.
================================================================================
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Optimize,
    [switch]$ComponentCleanup,
    [char]$DriveLetter=$env:SystemDrive.TrimEnd(':')[0],
    [string]$LogRoot="$env:ProgramData\WindowsDiskOptimization\Logs"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference='Stop'
$runPath=Join-Path $LogRoot (Get-Date -Format 'yyyyMMdd_HHmmss')
$warnings=New-Object System.Collections.Generic.List[string]

function Test-Admin{
    $id=[Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

try{
    if($env:OS -ne 'Windows_NT'){throw 'Windows is required.'}
    if(($Optimize -or $ComponentCleanup) -and -not(Test-Admin)){throw 'Run PowerShell as Administrator for optimisation.'}
    New-Item $runPath -ItemType Directory -Force|Out-Null

    Get-Volume -ErrorAction SilentlyContinue|
        Select-Object DriveLetter,FileSystemLabel,FileSystem,DriveType,HealthStatus,Size,SizeRemaining|
        Export-Csv (Join-Path $runPath 'Volumes.csv') -NoTypeInformation
    Get-PhysicalDisk -ErrorAction SilentlyContinue|
        Select-Object FriendlyName,SerialNumber,MediaType,HealthStatus,OperationalStatus,Size|
        Export-Csv (Join-Path $runPath 'PhysicalDisks.csv') -NoTypeInformation
    fsutil.exe behavior query DisableDeleteNotify 2>&1|Out-File (Join-Path $runPath 'TrimStatus.txt')
    dism.exe /Online /Cleanup-Image /AnalyzeComponentStore 2>&1|Out-File (Join-Path $runPath 'ComponentStoreAnalysis.txt')

    if($Optimize -and $PSCmdlet.ShouldProcess("Drive $DriveLetter",'Run Windows volume optimisation')){
        Optimize-Volume -DriveLetter $DriveLetter -Verbose 4>&1|Out-File (Join-Path $runPath 'OptimizeVolume.txt')
    }

    if($ComponentCleanup -and $PSCmdlet.ShouldProcess('Windows component store','Run component cleanup')){
        dism.exe /Online /Cleanup-Image /StartComponentCleanup /NoRestart 2>&1|
            Out-File (Join-Path $runPath 'ComponentCleanup.txt')
        if($LASTEXITCODE -notin 0,3010){$warnings.Add("Component cleanup returned $LASTEXITCODE")}
    }

    $warnings|Out-File (Join-Path $runPath 'Warnings.txt')
    if($warnings.Count -gt 0){Write-Host "[WARN] Completed with warnings. Logs: $runPath" -ForegroundColor Yellow;exit 2}
    Write-Host "[OK] Completed. Logs: $runPath" -ForegroundColor Green;exit 0
}catch{Write-Error $_.Exception.Message;exit 1}