<#
==========================================================================
  AUTHOR: Michael Dziegiel
  DATE  : 2023/12/13

  DESCRIPTION Create GIDAdmin account 
              Sets a Random password
              Sets password to not expire
 
==========================================================================
#>
function New-RandomPassword {
    param(
        [Parameter()][int]$PasswordLength = 15,
        [Parameter()][int]$NumberOfNonAlphaNumericCharacters = 5
        #[Parameter()][switch]$ConvertToSecureString
    )
    #Genereate Random Password
    
    Add-Type -AssemblyName 'System.Web'
    $password = [System.Web.Security.Membership]::GeneratePassword($Passwordlength,$NumberOfNonAlphaNumericCharacters)
    Return $password
   }

$userName = "GIDAdmin"
$userexist = (Get-LocalUser).Name -Contains $userName
if($userexist -eq $false) {
  try{ 
     [String]$PWD = New-RandomPassword -PasswordLength 16 -NumberOfNonAlphaNumericCharacters 8
        New-LocalUser -Name $userName -Password (ConvertTo-SecureString -String $PWD -AsPlainText -Force) -FullName $userName
        Set-LocalUser -Name $userName -PasswordNeverExpires
     Exit 0
   }   
  Catch {
     Write-error $_
     Exit 1
   }
}