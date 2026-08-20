<#
.AUTHOR      : Michael Dziegiel
.SCRIPT      : CleanupDisabledUsers
.SYNOPSIS    : Stage and delete disabled Active Directory user accounts.
.DESCRIPTION : Lists disabled accounts, sends notifications, and can delete accounts when requested.
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------
[CmdletBinding()]
PARAM(	
	[parameter(ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				Mandatory=$true)]
	[Alias('days')]
	[string]$TimeSpan,
	[parameter(ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				Mandatory=$false)]
	[string]$SearchBase,
	[parameter(ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				Mandatory=$false)]
	[switch]$SendMail=$true,
	[parameter(ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				Mandatory=$false)]
	[switch]$Notify=$false,
	[parameter(ValueFromPipeline=$true,
				ValueFromPipelineByPropertyName=$true,
				Mandatory=$false)]
	[switch]$whatIf=$false
)

BEGIN
{
	#Set mail details
	$SMTP = @{} 
	$SMTP.Address = "SMTP.DOMAIN.COM"
	$SMTP.From = "Servicedesk <Servicedesk@domain.com>"
	$SMTP.To = "someone@domain.com"
	$SMTP.NotificationTemplate = "MailTemplateAccountRemovalNotification.html"
	$SMTP.RemovedAccountsTemplate = "MailTemplateRemovedAccounts.html"
}
#-----------------------------------------------------------[Script]------------------------------------------------------------
PROCESS
{
	#Send notification one week before with accounts that are going to be removed
	if ($Notify)
	{
		#Set timespan - 7 days so we can send the notification one week before the actual removal of the accounts
		$TimeSpan = ($TimeSpan - 7)

		#Get the users accounts that will be deleted within x days that are already disabled and are not on hold
		$Accounts = Search-ADAccount -AccountInactive -TimeSpan "$TimeSpan" -UsersOnly -SearchBase $SearchBase | `
		Get-ADUser -Properties description,lastlogondate | `
		Where-Object { $_.Enabled -eq $false -and $_.Description -notlike '*on hold*'} | `
		Select name,lastlogondate
		
		#Are there accounts that will be deleted?
		if ($Accounts.Length -gt 0) 
		{
			#Convert results to HTML Table
			$Table = $Accounts | ConvertTo-Html -Fragment

			#Create the mail template
			$SMTP.Subject = "Listed User accounts will be removed in 7 days"

			$mailTemplate = (Get-Content ($PSScriptRoot + '\' + $SMTP.NotificationTemplate)) | ForEach-Object {
				$_ 	-replace '{{amount}}', $Accounts.Length `
				-replace '{{Table}}', $Table `
				-replace '{{TimeSpan}}', $TimeSpan `
			} | Out-String	

			#Send notification mail
			send-MailMessage -SmtpServer $SMTP.address -To $SMTP.To-From $SMTP.From -Subject $SMTP.Subject -Body $mailTemplate -BodyAsHtml -Priority High
		}
	}
	else
	{
		#Get the accounts to delete, we get the first so we can send an email with the deleted accounts
		$Accounts = Search-ADAccount -AccountInactive -TimeSpan "$TimeSpan" -UsersOnly -SearchBase $SearchBase | `
		Get-ADUser -Properties description |
		Where-Object { $_.Enabled -eq $false -and $_.Description -notlike '*on hold*'} | `
		Select Name,DistinguishedName
		
		#Are there accounts that will be deleted?
		if ($Accounts.Length -gt 0) 
		{
			Foreach ($User in $Accounts)
			{
				#Remove the user account
				#Using the Remove-ADObject to delete any device that's connected to exchange within this user account
				Get-ADUser -Identity $User.DistinguishedName | Remove-ADObject -Recursive -Confirm:$false -WhatIf:$whatIf
			}

			#Convert results to HTML Table
			$Table = $Accounts | ConvertTo-Html -Fragment

			#Create the mail template
			if ($whatIf)
			{
				$SMTP.Subject = "[DEMOMODUS] User accounts removed"
			}
			else
			{
				$SMTP.Subject = "User accounts removed"
			}
			
			$mailTemplate = (Get-Content ($PSScriptRoot + '\' + $SMTP.RemovedAccountsTemplate)) | ForEach-Object {
				$_ 	-replace '{{amount}}', $Accounts.Length `
				-replace '{{Table}}', $Table `
			} | Out-String	

			#Send notification mail
			send-MailMessage -SmtpServer $SMTP.address -To $SMTP.To-From $SMTP.From -Subject $SMTP.Subject -Body $mailTemplate -BodyAsHtml
		}
	}
}