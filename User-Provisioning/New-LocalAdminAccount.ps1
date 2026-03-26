<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2023/12/13

  Script: New-LocalAdminAccount.ps1

  Scope: Creates new local admin account and adds to Administrators group

  NOTE : Replace placeholders or inject securely (Intune / env vars)
==========================================================================
#>

# 🔒 DO NOT HARDCODE IN PRODUCTION
$Username = "<LOCAL_ADMIN_USERNAME>"
$Password = "<LOCAL_ADMIN_PASSWORD>"

$group = "Administrators"

$adsidetails = [ADSI]"WinNT://$env:COMPUTERNAME"

$usercheck = $adsidetails.Children | Where-Object {
    $_.SchemaClassName -eq 'user' -and $_.Name -eq $Username
}

if ($usercheck -eq $null)
{
    Write-Host "Creating the new local user $Username."

    & NET USER $Username $Password /add /y /expires:never
    
    Write-Host "Adding local user $Username to the $group."

    & NET LOCALGROUP $group $Username /add
}
else
{
    Write-Host "Setting the password for existing local user $Username."

    $user = [ADSI]"WinNT://$env:COMPUTERNAME/$Username,user"
    $user.SetPassword($Password)
}

Write-Host "Setting the password for $Username to never expire."

& WMIC USERACCOUNT WHERE "Name='$Username'" SET PasswordExpires=FALSE
