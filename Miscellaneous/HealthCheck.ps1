<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : HealthCheck
SYNOPSIS    : Checks Windows server health across hardware, software, and
              network basics.
DESCRIPTION : Checks Windows server health across hardware, software, and
              network basics.
================================================================================
#>
Write-Host "===================================" -ForegroundColor Cyan
Write-Host " Windows Server Health Check"
Write-Host "===================================" -ForegroundColor Cyan

# Computer Name
Write-Host "`nComputer Name:"
hostname

# Operating System
Write-Host "`nOperating System:"
(Get-CimInstance Win32_OperatingSystem).Caption

# Last Boot Time
Write-Host "`nLast Boot Time:"
(Get-CimInstance Win32_OperatingSystem).LastBootUpTime

# CPU
Write-Host "`nCPU:"
(Get-CimInstance Win32_Processor).Name

# Memory
Write-Host "`nMemory:"
Get-CimInstance Win32_OperatingSystem |
Select-Object @{
    Name="Total RAM (GB)"
    Expression={[math]::Round($_.TotalVisibleMemorySize/1MB,2)}
}, @{
    Name="Free RAM (GB)"
    Expression={[math]::Round($_.FreePhysicalMemory/1MB,2)}
}

# Disk Space
Write-Host "`nDisk Usage:"
Get-PSDrive -PSProvider FileSystem |
Select-Object Name,
@{Name="Used (GB)";Expression={[math]::Round(($_.Used)/1GB,2)}},
@{Name="Free (GB)";Expression={[math]::Round($_.Free/1GB,2)}}

# Top Running Services
Write-Host "`nRunning Services:"
Get-Service |
Where-Object {$_.Status -eq "Running"} |
Select-Object -First 10 Name,DisplayName

Write-Host "`nHealth Check Completed Successfully." -ForegroundColor Green