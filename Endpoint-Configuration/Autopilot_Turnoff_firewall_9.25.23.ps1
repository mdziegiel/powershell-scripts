$ErrorActionPreference = "silentlycontinue"
$User = New-Object System.Security.Principal.NTAccount((Get-WmiObject -Class win32_computersystem).UserName)

Write-Host "Current logged on user is: " $User.value

 

 

if ($User.value -like "$env:COMPUTERNAME\*")

{

   Write-Host "A local user logged on, assume there is no AD connectivity and exit"

   Stop-Transcript

exit 0

}

else

{

   $RS = Test-Path 'C:\Program Files\GID\logs' -PathType Container
   If ($RS){
   Write-Host "GID log present"
   }Else{
   
   Stop-Transcript
   EXIT 99
       
   }
 

   #Do Stuff here

}

   Stop-Transcript