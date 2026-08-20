<#
.AUTHOR      : Michael Dziegiel
.SCRIPT      : CleanupInActiveComputers
.SYNOPSIS    : Stage and disable inactive Active Directory computer accounts.
.DESCRIPTION : Lists inactive computer accounts, sends notifications, and can disable them when requested.
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
	$SMTP.NotificationTemplate = "MailTemplateDisabledComputersNotification.html"
	$SMTP.RemovedAccountsTemplate = "MailTemplateRemovedAccounts.html"
}
#-----------------------------------------------------------[Script]------------------------------------------------------------
PROCESS
{
	#Send notification one week before with computer accounts that are going to be disabled
	if ($Notify)
	{
		#Set timespan - 7 days so we can send the notification one week before the actual removal of the accounts
		$TimeSpan = ($TimeSpan - 7)

		#Get the computers accounts that will be deleted within x days that are not on hold
		$Computers = Search-ADAccount -AccountInactive -TimeSpan "$TimeSpan" -ComputersOnly -SearchBase $SearchBase | `
		Get-ADComputer -Properties description,lastlogondate | `
		Where-Object { $_.Description -notlike '*on hold*' } | `
		Select name,lastlogondate
		
		#Are there computers that will be disabled?
		if ($Computers.Length -gt 0) 
		{
			#Convert results to HTML Table
			$Table = $Computers | ConvertTo-Html -Fragment

			#Create the mail template
			$SMTP.Subject = "Listed Computers objects will be disabled in 7 days"

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
		#Get the computers account to disable, we get them first so we can send an email with the disabled accounts
		$Computers = Search-ADAccount -AccountInactive -TimeSpan "$TimeSpan" -ComputersOnly -SearchBase $SearchBase | `
		Get-ADComputer -Properties description,lastlogondate | `
		Where-Object { $_.Description -notlike '*on hold*'} | `
		Select Name,lastlogondate
		
		#Are there computers that will be disabled?
		if ($Computers.Length -gt 0)
		{
			Foreach ($computer in $Computers)
			{
				#Disable the computer account
				Disable-ADAccount -Identity $Computer.DistinguishedName -Confirm:$false -WhatIf:$whatIf
			}

			#Convert results to HTML Table
			$Table = $Accounts | ConvertTo-Html -Fragment

			#Create the mail template
			if ($whatIf)
			{
				$SMTP.Subject = "[DEMOMODUS] Computer accounts disabled"
			}
			else
			{
				$SMTP.Subject = "Computer accounts disabled"
			}
			
			$mailTemplate = (Get-Content ($PSScriptRoot + '\' + $SMTP.DisabledComputersTemplate)) | ForEach-Object {
				$_ 	-replace '{{amount}}', $Computers.Length `
				-replace '{{Table}}', $Table `
			} | Out-String	

			#Send notification mail
			send-MailMessage -SmtpServer $SMTP.address -To $SMTP.To-From $SMTP.From -Subject $SMTP.Subject -Body $mailTemplate -BodyAsHtml
		}
	}
}