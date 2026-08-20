<#
.AUTHOR      : Michael Dziegiel
.SCRIPT      : Add-UsersToGroup
.SYNOPSIS    : Bulk add users from a CSV file to an Active Directory group.
.DESCRIPTION : Import CSV file and add each user to the target group.
#>

[CmdletBinding()]
param (
    [Parameter(
      Mandatory = $true,
      HelpMessage = "Group name"
    )]
    [string] $GroupName = "",

    [Parameter(
      Mandatory = $true,
      HelpMessage = "Path to CSV file"
    )]
    [string] $Path = "",

    [Parameter(
      Mandatory = $false,
      HelpMessage = "CSV file delimiter"
    )]
    [string] $Delimiter = ",",

    [Parameter(
      Mandatory = $false,
      HelpMessage = "Find users on DisplayName, Email or UserPrincipalName"
    )]
    [ValidateSet("DisplayName", "Email", "UserPrincipalName")]
    [string] $Filter = "DisplayName"
)

Function Add-UsersToGroup {
    <#
    .SYNOPSIS
      Get users from the requested DN
    #>
    process{
        # Import the CSV File
        $users = (Import-Csv -Path $path -Delimiter $delimiter -header "name").name

        # Find the users in the Active Directory
        $users | ForEach {
            $user =  Get-ADUser -filter "$filter -eq '$_'" | Select ObjectGUID 

            if ($user) {
                Add-ADGroupMember -Identity $groupName -Members $user
                Write-Host "$_ added to the group"
            }else {
                Write-Warning "$_ not found in the Active Directory"
            }
        }
    }
}

# Load the Active Directory Module
Import-Module -Name ActiveDirectory

# Add user from CSV to given Group
Add-UsersToGroup