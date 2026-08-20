<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get-MailboxFolderPermissions
SYNOPSIS    : Enumerates mailbox folders in a mailbox and reports users with
              explicit non-default, non-anonymous permissions on
              each folder.
DESCRIPTION : Enumerates mailbox folders in a mailbox and reports users with
              explicit non-default, non-anonymous permissions on
              each folder.
================================================================================
#>
#Requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$MailboxIdentity = '<MailboxIdentity>'
)

# Script to list mailbox folder permissions, who has access on what folders etc.
$permissions = @()
$Folders = Get-MailboxFolderStatistics $MailboxIdentity | % {$_.folderpath} | % { $_.replace('/', [char]92) }
$list = ForEach ($F in $Folders) {
    $FolderKey = $MailboxIdentity + ":" + $F
    $Permissions += Get-MailboxFolderPermission -identity $FolderKey -ErrorAction SilentlyContinue | Where-Object {$_.User -notlike 'Default' -and $_.User -notlike 'Anonymous' -and $_.AccessRights -notlike 'None'}
}
$permissions