<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Install-Firefox
SYNOPSIS    : Winget wrapper to install Mozilla Firefox
DESCRIPTION : Winget wrapper to install Mozilla Firefox
================================================================================
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$PackageId = 'Mozilla.Firefox'
$Arguments = @(
    'install',
    '--id', $PackageId,
    '--exact',
    '--silent',
    '--accept-package-agreements',
    '--accept-source-agreements'
)
if ($Force) { $Arguments += '--force' }

Write-Host "Installing Mozilla Firefox with winget..." -ForegroundColor Cyan
& winget @Arguments
exit $LASTEXITCODE