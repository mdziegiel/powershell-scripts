<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Remove Local Accounts
SYNOPSIS    : Removes non-approved local user accounts
DESCRIPTION : Enumerates all local user accounts and deletes any account
              not included in an approved allow list. Intended for
              enforcing a standardized local account baseline and
              reducing unauthorized access
==============================================================================
#>

$UserList = @("Administrator","GIDAdmin","DefaultAccount","Guest","IEUser","sshd","WDAGUtilityAccount")
Get-WMIObject -Class Win32_UserAccount | Where-Object {$_.Name -notin $UserList} | Foreach {net user $_.Name /delete}
