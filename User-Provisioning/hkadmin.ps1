<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2025/04/10

  Script: HKAdmin.ps1

  Scope: Creates an Admin Account and sets the standard password to never expire.
==========================================================================
#>

# Define account credentials
$Username = "hkadmin"
$PasswordPlain = "BaconCat30"
$Password = $PasswordPlain | ConvertTo-SecureString -AsPlainText -Force

# Check if the user exists
if (-Not (Get-LocalUser -Name $Username -ErrorAction SilentlyContinue)) {
    try {
        # Create the user
        New-LocalUser -Name $Username -Password $Password -FullName "HK Admin" -Description "Admin created by Intune"

        # Set password to never expire
        Set-LocalUser -Name $Username -PasswordNeverExpires $true

        # Add to Administrators group
        Add-LocalGroupMember -Group "Administrators" -Member $Username

        Write-Output "User '$Username' created and added to Administrators group."
        Exit 0
    }
    catch {
        Write-Error "Failed to create user '$Username': $_"
        Exit 1
    }
}
else {
    Write-Output "User '$Username' already exists. No action taken."
    Exit 0
}
