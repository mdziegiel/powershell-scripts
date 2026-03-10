<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2023/12/13

  Script: New-Localadminaccount.ps1

  Scope: Creates new local admin account and add to Administrators group
 
==========================================================================
#>


$Username = "gidadmin" 

$Password = "G!DT3mpAdm1n$2023" 

$group = "Administrators"

$adsidetails = [ADSI]"WinNT://$env:COMPUTERNAME"

$usercheck = $adsidetails.Children | where {$_.SchemaClassName -eq 'user' -and $_.Name -eq $Username }

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

    $existing.SetPassword($Password)
}

Write-Host "Setting the password for $Username never expires."

& WMIC USERACCOUNT WHERE "Name='$Username'" SET PasswordExpires=FALSE