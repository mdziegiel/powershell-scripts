<#
.AUTHOR      : Michael Dziegiel
.SCRIPT      : Get-ADComputers
.SYNOPSIS    : Export Active Directory computer objects and properties to CSV.
.DESCRIPTION : Collects enabled computer objects and writes the results to CSV.
#>

param(
  [Parameter(
    Mandatory = $false,
    HelpMessage = "Enter the searchbase between quotes or multiple separated with a comma"
    )]
  [string[]]$searchBase,

  [Parameter(
    Mandatory = $false,
    HelpMessage = "Get computers that are enabled, disabled or both"
  )]
    [ValidateSet("true", "false", "both")]
  [string]$enabled = "true",

  [Parameter(
    Mandatory = $false,
    HelpMessage = "Enter path to save the CSV file"
  )]
  [string]$CSVpath
)

Function Get-Computers {
    <#
    .SYNOPSIS
      Get computers from the requested DN
    #>
    param(
      [Parameter(Mandatory = $true)]
      $dn
    )
    process{
      # Set the properties to retrieve
      $properties = @(
        'Name',
        'CanonicalName',
        'OperatingSystem',
        'OperatingSystemVersion',
        'LastLogonDate',
        'LogonCount',
        'BadLogonCount',
        'IPv4Address',
        'Enabled',
        'whenCreated'
      )

      # Get enabled, disabled or both computers
      switch ($enabled)
      {
        "true" {$filter = "enabled -eq 'true'"}
        "false" {$filter = "enabled -eq 'false'"}
        "both" {$filter = "*"}
      }

      # Get the computers
      Get-ADComputer -Filter $filter -searchBase $dn -Properties $properties | Select-Object $properties
    }
}


Function Get-AllADComputers {
  <#
    .SYNOPSIS
      Get all AD computers
  #>
  process {
    Write-Host "Collecting computers" -ForegroundColor Cyan
    $computers = @()

    # Collect computers
    if ($searchBase) {
      # Get the requested mailboxes
      foreach ($dn in $searchBase) {
        Write-Host "- Get computers in $dn" -ForegroundColor Cyan
        $computers += Get-Computers -dn $dn
      }
      }else{
        # Get distinguishedName of the domain
        $dn = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
        Write-Host "- Get computers in $dn" -ForegroundColor Cyan
        $computers += Get-Computers -dn $dn
      }

    # Loop through all computers
    $computers | ForEach-Object {

      [pscustomobject]@{
        "Name" = $_.Name
        "CanonicalName" = $_.CanonicalName
        "OS" = $_.OperatingSystem
        "OS Version" = $_.OperatingSystemVersion
        "Last Logon" = $_.lastLogonDate
        "Logon Count" = $_.logonCount
        "Bad Logon Count" = $_.BadLogonCount
        "IP Address" = $_.IPv4Address
        "Mobile" = $_.mobile
        "Enabled" = if ($_.Enabled) {"enabled"} else {"disabled"}
        "Date created" = $_.whenCreated
      }
    }
  }
}

If ($CSVpath) {
  # Get all AD Computers
  Get-AllADComputers | Sort-Object Name | Export-CSV -Path $CSVpath -NoTypeInformation -Encoding UTF8
  if ((Get-Item $CSVpath).Length -gt 0) {
    Write-Host "Report finished and saved in $CSVpath" -ForegroundColor Green
  
    # Open the CSV file
    Invoke-Item $CSVpath
  
  }else{
    Write-Host "Failed to create report" -ForegroundColor Red
  }
}
Else {
  Get-AllADComputers | Sort-Object Name
}