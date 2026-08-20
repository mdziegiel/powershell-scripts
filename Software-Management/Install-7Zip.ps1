<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Install-7Zip
SYNOPSIS    : Winget wrapper to install 7-Zip
DESCRIPTION : Winget wrapper to install 7-Zip
================================================================================
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$PackageId = '7zip.7zip'
$Arguments = @(
    'install',
    '--id', $PackageId,
    '--exact',
    '--silent',
    '--accept-package-agreements',
    '--accept-source-agreements'
)
if ($Force) { $Arguments += '--force' }

Write-Host "Installing 7-Zip with winget..." -ForegroundColor Cyan
& winget @Arguments
exit $LASTEXITCODE