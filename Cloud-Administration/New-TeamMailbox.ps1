<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : New-TeamMailbox
SYNOPSIS    : This scripts creates a new shared mailbox (aka team mailbox) and
              security groups for full access and and send-as
              delegation
DESCRIPTION : This scripts creates a new shared mailbox (aka team mailbox) and
              security groups for full access and and send-as
              delegation. Security groups are created using a
              naming convention.
================================================================================
#>
#Requires -Modules ExchangeOnlineManagement, MicrosoftTeams, ActiveDirectory
param (
  [parameter(Mandatory,HelpMessage='Team Mailbox Name')]
  [string]$TeamMailboxName,
  [parameter(Mandatory,HelpMessage='Team Mailbox Display Name')]
  [string]$TeamMailboxDisplayName,
  [parameter(Mandatory,HelpMessage='Team Mailbox Alias')]
  [string]$TeamMailboxAlias,
  [string]$TeamMailboxSmtpAddress = '',
  [string]$DepartmentPrefix = '',
  $GroupFullAccessMembers = @(''),
  $GroupSendAsMember = @()
)

# Script Path
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Path

if(Test-Path -Path ('{0}\Settings.xml' -f $scriptPath)) {
    # Load Script settings
    [xml]$Config = Get-Content -Path ('{0}\Settings.xml' -f $scriptPath)
    
    Write-Verbose -Message 'Loading script settings'
    
    # Group settings
    $groupPrefix = $Config.Settings.GroupSettings.Prefix
    $groupSendAsSuffix = $Config.Settings.GroupSettings.SendAsSuffix
    $groupFullAccessSuffix = $Config.Settings.GroupSettings.FullAccessSuffix
    $groupTargetOU = $Config.Settings.GroupSettings.TargetOU
    $groupDomain = $Config.Settings.GroupSettings.Domain
    $groupPrefixSeperator = $Config.Settings.GroupSettings.Seperator
    
    # Team mailbox settings
    $teamMailboxTargetOU = $Config.Settings.AccountSettings.TargetOU

    # General settings
    $sleepSeconds = $Config.Settings.GeneralSettings.Sleep

    Write-Verbose -Message 'Script settings loaded'    
}
else {
    Write-Error -Message 'Script settings file settings.xml missing'
    exit 99
}

# Add department prefix to group prefix, if configured
if($DepartmentPrefix -ne '') {
    # Change pattern as needed
    $groupPrefix = ('{0}{1}{2}' -f $groupPrefix, $DepartmentPrefix, $groupPrefixSeperator)
}

# Create shared team mailbox
Write-Verbose -Message ('New-Mailbox -Shared -Name {0} -Alias {1}' -f $TeamMailboxName, $TeamMailboxAlias)

if ($TeamMailboxSmtpAddress -ne '') 
{
    $null = New-Mailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -PrimarySmtpAddress $TeamMailboxSmtpAddress -DisplayName $TeamMailboxDisplayName
}
else
{
    $null = New-Mailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -DisplayName $TeamMailboxDisplayName
}

# Create Full Access group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $TeamMailboxAlias, $groupFullAccessSuffix)
$notes = ('FullAccess for mailbox: {0}' -f $TeamMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

Write-Host ('Creating new FullAccess Group: {0}' -f $groupName)

Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

if(($GroupFullAccessMembers | Measure-Object).Count -ne 0) {

    Write-Host ('Creating FullAccess group and adding members: {0}' -f $groupName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupFullAccessMembers -Notes $notes

    Start-Sleep -Seconds $sleepSeconds

    # Hide FullAccess group from GAL
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {

    Write-Host ('Creating empty FullAccess group: {0}' -f $groupName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes

    Start-Sleep -Seconds $sleepSeconds

    # Hide FullAccess group from GAL    
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# Add full access group to mailbox permissions

Write-Verbose -Message ('Add-MailboxPermission -Identity {0} -User {1}' -f $TeamMailboxName, $primaryEmail)

$null = Add-MailboxPermission -Identity $TeamMailboxName -User $primaryEmail -AccessRights FullAccess -InheritanceType all

# Create Send As group
$groupName = ('{0}{1}{2}' -f $groupPrefix, $TeamMailboxAlias, $groupSendAsSuffix)
$notes = ('SendAs for mailbox: {0}' -f $TeamMailboxName)
$primaryEmail = ('{0}@{1}' -f $groupName, $groupDomain)

Write-Host ('Creating new SendAs Group: {0}' -f $groupName)

Write-Verbose -Message ('New-DistributionGroup -Name {0} -Type Security -OrganizationalUnit {1} -PrimarySmtpAddress {2}' -f $groupName, $groupTargetOU, $primaryEmail)

if(($GroupSendAsMember | Measure-Object).Count -ne 0) {

    Write-Host ('Creating SendAs group and adding members: {0}' -f $groupName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupSendAsMember -Notes $notes

    Start-Sleep -Seconds $sleepSeconds

    # Hide SendAs from GAL
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {

    Write-Host ('Not empty SendAs group: {0}' -f $groupName)

    $null = New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes

    Start-Sleep -Seconds $sleepSeconds

    # Hide SendAs from GAL
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# Add SendAs permission
Write-Verbose -Message ('Add-ADPermission -Identity {0} -User {1}' -f $TeamMailboxName, $groupName)

$null = Add-ADPermission -Identity $TeamMailboxName -User $groupName -ExtendedRights 'Send-As'

Write-Host ('Script finished. Team mailbox {0} created.' -f $TeamMailboxName)