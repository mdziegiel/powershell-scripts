<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Teams-Phone-User-Assignment-Phase1
SYNOPSIS    : Assign voice-related Teams policies to users
DESCRIPTION : Assign voice-related Teams policies to users
================================================================================
#>
#Requires -Modules MicrosoftTeams
Write-Host ""
Write-Host "This script assigns to users  different voice-related policies in Teams as part of the phase 1 (Preparation). Settings are defined in the CSV file"
Write-Host "Make sure you get the latest versions for MicrosoftTeams module and PowerShell."
Write-Host ""

# Establish Teams remote PowerShell session
Write-Host ""
Write-Host "Connecting to Microsoft Teams AC"
Write-Host ""
$Session = Connect-MicrosoftTeams

#GLOBAL VARIABLES
[String]$date = Get-Date -UFormat "%Y_%m_%d_%H_%M_%S"
$DomainName = (Get-CsTenant).DisplayName
$TargetPath = 'C:\Temp\'+$DomainName+'\Teams-Output-Config'+"\v"+$date+"\"                                                    
$TestPath = test-path -path $TargetPath
if ($TestPath -ne $true) {New-Item -ItemType directory -Path $TargetPath | Out-Null
    $MessagePath = 'Creating directory to write file to '+$TargetPath+'.Your files will be uploaded in this folder'
    write-Host  $MessagePath}
else {
    $MessagePath = 'Your files will be uploaded to '+$TargetPath
    Write-Host $MessagePath}
[String]$transcriptFilePath = $TargetPath+"Transcript_TeamsVoicePolicies_UserProvisioning_Phase1_$date.txt"
[String]$logFilePath = $TargetPath+"Log_TeamsVoicePolicies_UserProvisioning_Phase1_$date.csv"
[String]$ExportUsersFilePath = $TargetPath+"ExportUsers_TeamsVoicePolicies_UserProvisioning_Phase1_$date.csv"
$users = $null
$global:errorCount = 0
$global:errorCountUser = 0
#Log Initialization
Add-Content -Path $logFilePath -Value "LogLevel,UPN,Parameter,Message"

#output file - Path 
Write-Host "-----------------------------------------"
Write-Host "TranscriptFile Path :"$transcriptFilePath
Write-Host "LogFile Path :"$logFilePath
Write-Host "UsersExport Path :"$ExportUsersFilePath
Write-Host "-----------------------------------------"

#Select CSV file dialog
Function Select-CSVFile
{
    param([string]$Title="Please select the batch file",[string]$Directory=$((Get-Location).Path + "\Batches"),[string]$Filter="Comma Seperated Values | *.csv")
	[System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") | Out-Null
	$objForm = New-Object Microsoft.Win32.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$show = $objForm.ShowDialog()
    
	If ($show -eq $true)
	{
		#Get file path
        [String]$csvFile = $objForm.FileName

		#Check file
        Write-Warning "The following file was selected: $($csvFile.Split("\")[$_.Length-1])"
		$continue = Read-Host "Press [C] to continue"

		If ($continue.ToLower() -eq "c")
		{
			Return $csvFile
		}
    }
    
    Return 0
}


#Log errors to file
Function Add-Log
{
    Param([String]$UPN, [ValidateSet("INFO", "WARNING", "ERROR")][string]$Level="ERROR", [String]$Parameter, [String]$Message)
    switch ($Level)
    {
        "INFO"
        {
        Write-Host "INFO: "$user.UserPrincipalName " : " $Message -ForegroundColor Green
        }
        "WARNING"
        {
        Write-Host "WARNING: "$user.UserPrincipalName " : " $Message -ForegroundColor Yellow
        }
        "ERROR"
        {
        Write-Host "ERROR: "$user.UserPrincipalName " : " $Message -ForegroundColor Red
        $global:errorCount++
        $global:errorCountUser++
        }
    }
    $log = $level+ ',' +$UPN + ',' + $Parameter + ',"' + $Message + '"'
    Add-Content -Path $logFilePath -Value $log
}

# Create Transcript file for analysis later when script is completed.  Checking for errors etc.
Start-Transcript -path $transcriptFilePath -append

Write-Host "Select the CSV file used for the batch" -ForegroundColor Yellow
$csvFile = Select-CSVFile
$users = Import-CSV $csvFile
$count = $users.count

write-host " "
write-host " "
write-host "From the CSV file, We have found "$count " Users to assign the relative policies" 
write-host "Processing "$count "Users...Please wait..." -foregroundcolor Yellow -backgroundcolor Black
write-host " "
write-host " "

# Run the PowerShell commands for each user, line by line in the CSV file.

ForEach ($user in $users) 
{
    $global:errorCountUser = 0
    
    $userEnabled = $( try{ Get-csOnlineUser -Identity $user.UserPrincipalName |select AccountEnabled} catch {$Null})
    if($userEnabled -ne $null){
     
       
    #checking if the user account is enabled or error in the log
    if($userEnabled.AccountEnabled -eq 'TRUE')
    {
        
        #Assign Calling Policy
	    If ($user.TeamsCallingPolicy -ne "")
          {
          Try { Grant-CsTeamsCallingPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsCallingPolicy -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Parameter "Teams Calling policy" -Message $Error[0] }
         }
       
       #Assign Caller ID Policy (CLID)
	    If ($user.CallingLineIdentity -ne "")
          {
          Try { Grant-CsCallingLineIdentity -Identity $user.UserPrincipalName -PolicyName $user.CallingLineIdentity -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Teams CLID policy" -Message $Error[0] }
         }

       #Assign MOH POlicy
       If ($user.TeamsMOHPolicy -ne "")
          {
          Try { Grant-CsTeamsCallHoldPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsMOHPolicy -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "MOH policy" -Message $Error[0] }
         }
             
	   #Assign Call Park Policy
       If ($user.TeamsCallParkPolicy -ne "")
          {
          Try { Grant-CsTeamsCallParkPolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsCallParkPolicy -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Call Park policy" -Message $Error[0] }
         }
    
       #Assign DialPlan Policy (usually seen as Normalization Rules)
       If ($user.TenantDialPlan -ne "")
          {
          Try { Grant-CsTenantdialplan -Identity $user.UserPrincipalName -PolicyName $user.TenantDialPlan -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Dial Plan" -Message $Error[0] }
         }
        
		#Assign IP Phone Policy if you have some phone devices in the spreadsheet (common area phones, hotdesking, or meeting room accounts)
        If ($user.TeamsIPPhonePolicy -ne "")
          {
          Try { Grant-CsTeamsIPPhonePolicy -Identity $user.UserPrincipalName -PolicyName $user.TeamsIPPhonePolicy -ErrorAction Stop }
         Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "IP Phone Policy" -Message $Error[0] }
         }

                        
        #Assign the VRP if VRP is filled out in the CSV - Make sure VRP means OperatorType = "DirectRouting"
        If ($user.OnlineVoiceRoutingPolicy -ne "")
        {
        Try { Grant-CsOnlineVoiceRoutingPolicy -Identity $user.UserPrincipalName -PolicyName $user.OnlineVoiceRoutingPolicy -ErrorAction Stop }
        Catch { Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Voice Routing Policy" -Message $Error[0] }
        }
       
       
    }else{
          Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Disabled" -Message " The user account is disabled"    
    }
          
  #For this user, review if any error has been reported for Logs
  If ($global:errorCountUser -ne 0)
    {
       Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "Errors during Assignment" -Message "Assignment has completed with errors, please see logs in log Folder."
    }else    
    {
       Add-Log -UPN $user.UserPrincipalName -Level INFO -Parameter "Assignment completed" -Message "Assignment has completed with no errors"
    }

 }else{ Add-Log -UPN $user.UserPrincipalName -Level ERROR -Parameter "UPN not found" -Message "The UPN can not be found"}

}

# Assignment completed
write-host ""
write-host ""
write-host "PROGRESS STATUS : Completed Voice policies assignment for this users batch !" -foregroundcolor White 
write-host ""    
    If ($global:errorCount -ne 0)
    {
       Write-Host "Assignment of the batch has completed with errors, please see transcript and logs in Logs Folder." -Foregroundcolor Yellow 
    }
    Else
    {
       Write-Host "Assignment of the batch has completed with no errors. Check UsersExport CSV file for any discrepancy." -ForegroundColor Green
    }


#Creating a user assignment export with new Users policies values  
write-host " "
Write-host "User Assignment Status Report being generated" -foregroundcolor White -backgroundcolor Black
write-host " "
ForEach ($user in $users) 
{ 
    try{ Get-CsOnlineUser -Identity $user.UserPrincipalName | Select UserPrincipalName, DisplayName, AccountEnabled, EnterpriseVoiceEnabled, UsageLocation, Country, DialPlan, TenantDialPlan, InterpretedUserType, TeamsCallingPolicy, CallingLineIdentity, OnlineVoicemailPolicy, TeamsCallHoldPolicy, TeamsCallParkPolicy, TeamsEmergencyCallingPolicy, TeamsIPPhonePolicy, OnlineVoiceRoutingPolicy, ProvisioningStamp, SubProvisioningStamp | Export-csv -path $ExportUsersFilePath -Append -NoTypeInformation }
    Catch { Add-Log -UPN $user.UserPrincipalName -Level Error -Parameter "UPN not found" -Message "The user will not be in the User Export CSV file - UPN not found" }
} 
    
Write-Host "Closing Teams admin session.....Done!" -foregroundcolor White
Write-Host ""
Stop-Transcript

Disconnect-MicrosoftTeams