<#
.AUTHOR      : Michael Dziegiel
.SCRIPT      : Get-UserFromOU
.SYNOPSIS    : Export Active Directory users from a specific OU to CSV.
.DESCRIPTION : Collects users and common attributes from an OU and exports them to CSV.
#>

param(
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Enter the searchbase between quotes or multiple separated with a comma"
        )]
        [string]$searchBase,
    
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Enter path to save the CSV file"
    )]
    [string]$path = ".\ADUsers-$((Get-Date -format "MMM-dd-yyyy").ToString()).csv"
)

Function Get-Users {
    <#
    .SYNOPSIS
      Get users from the requested DN
    #>
    param(
      [Parameter(
        Mandatory = $true
      )]
      $dn
    )
    process{
      # Set the properties to retrieve
      $properties = @(
        'name',
        'userprincipalname',
        'mail',
        'title',
        'telephoneNumber',
        'mobile',
        'department',
        'extensionAttribute5',
        'extensionAttribute3',
        'extensionAttribute4'
      )

      # Get the user
      Get-ADUser -Filter "Enabled -eq 'true'" -searchBase $dn -properties $properties | where {$_.extensionAttribute5 -eq 'ListInDigitalReception'} | select $properties
    }
}


Function Get-AllADUsers {
    <#
      .SYNOPSIS
        Get all AD users
    #>
    process {
      Write-Host "Collecting users" -ForegroundColor Cyan
      $users = @()
  
      # Collect users
      $users += Get-Users -dn $searchBase
  
      # Loop through all users
      $users | ForEach {
  
        [pscustomobject]@{
          "Name" = $_.Name
          "UserPrincipalName" = $_.UserPrincipalName
          "Emailaddress" = $_.mail
          "Phone" = $_.telephoneNumber
          "Mobile" = $_.mobile
          "Job title" = $_.Title
          "Department" = $_.Department
          "Extension3" = $_.extensionAttribute3
          "Extension4" = $_.extensionAttribute4
        }
      }
    }
  }
  

Get-AllADUsers | Sort-Object Name | Export-CSV -Path $path -NoTypeInformation

if ((Get-Item $path).Length -gt 0) {
  Write-Host "Report finished and saved in $path" -ForegroundColor Green

  # Open the CSV file
  Invoke-Item $path

}else{
  Write-Host "Failed to create report" -ForegroundColor Red
}