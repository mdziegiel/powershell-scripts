<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Align Windows 11 Taskbar Left
SYNOPSIS    : Sets Windows 11 taskbar alignment to left
DESCRIPTION : Configures the registry value 'TaskbarAl' under
              HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced
              to 0, which aligns the taskbar to the left for the current user
==============================================================================
#>

$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$Al = "TaskbarAl"   # Taskbar alignment
$value = "0"

New-ItemProperty -Path $registryPath -Name $Al -Value $value -PropertyType DWORD -Force -ErrorAction Ignore
