<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Install-AdobeReader
SYNOPSIS    : Winget wrapper to install Adobe Acrobat Reader
DESCRIPTION : Winget wrapper to install Adobe Acrobat Reader
================================================================================
#>
[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$PackageId = 'Adobe.Acrobat.Reader.64-bit'
$Arguments = @(
    'install',
    '--id', $PackageId,
    '--exact',
    '--silent',
    '--accept-package-agreements',
    '--accept-source-agreements'
)
if ($Force) { $Arguments += '--force' }

Write-Host "Installing Adobe Acrobat Reader with winget..." -ForegroundColor Cyan
& winget @Arguments
exit $LASTEXITCODE