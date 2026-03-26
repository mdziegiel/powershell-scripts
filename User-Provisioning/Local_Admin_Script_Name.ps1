<#
==============================================================================
AUTHOR      : Michael Dziegiel
DATE        : 2025/04/10
SCRIPT      : <LOCAL_ADMIN_SCRIPT_NAME>.ps1
SYNOPSIS    : Creates a local administrator account
DESCRIPTION : Creates a local administrator account, sets the password to
              never expire, and adds the account to the local
              Administrators group

ORGANIZATION: <ORGANIZATION_NAME>

NOTES       : This script has been sanitized for public/shared use.
              Replace the placeholders below with values for your environment:

              <LOCAL_ADMIN_SCRIPT_NAME> = Script file name
              <LOCAL_ADMIN_USERNAME>    = Local admin account name
              <LOCAL_ADMIN_PASSWORD>    = Local admin password
              <LOCAL_ADMIN_FULLNAME>    = Display name for the account
              <ACCOUNT_DESCRIPTION>     = Account description
==============================================================================
#>

# Define account credentials
$Username      = "<LOCAL_ADMIN_USERNAME>"
$PasswordPlain = "<LOCAL_ADMIN_PASSWORD>"
$Password      = $PasswordPlain | ConvertTo-SecureString -AsPlainText -Force

# Check if the user exists
if (-not (Get-LocalUser -Name $Username -ErrorAction SilentlyContinue)) {
    try {
        # Create the user
        New-LocalUser -Name $Username -Password $Password -FullName "<LOCAL_ADMIN_FULLNAME>" -Description "<ACCOUNT_DESCRIPTION>"

        # Set password to never expire
        Set-LocalUser -Name $Username -PasswordNeverExpires $true

        # Add to Administrators group
        Add-LocalGroupMember -Group "Administrators" -Member $Username

        Write-Output "User '$Username' created and added to Administrators group."
        exit 0
    }
    catch {
        Write-Error "Failed to create user '$Username': $_"
        exit 1
    }
}
else {
    Write-Output "User '$Username' already exists. No action taken."
    exit 0
}
