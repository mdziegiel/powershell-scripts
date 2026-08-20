<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Install-Chrome
SYNOPSIS    : Winget wrapper to install Google Chrome
DESCRIPTION : Winget wrapper to install Google Chrome
================================================================================
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$PackageId = 'Google.Chrome'
$Arguments = @(
    'install',
    '--id', $PackageId,
    '--exact',
    '--silent',
    '--accept-package-agreements',
    '--accept-source-agreements'
)
if ($Force) { $Arguments += '--force' }

Write-Host "Installing Google Chrome with winget..." -ForegroundColor Cyan
& winget @Arguments
exit $LASTEXITCODE