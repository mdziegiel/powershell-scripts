<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : List-AllMembersOfADistributionList
SYNOPSIS    : Connects to Exchange Online, expands a distribution list, and
              prints each member’s display name and primary SMTP
              address.
DESCRIPTION : Connects to Exchange Online, expands a distribution list, and
              prints each member’s display name and primary SMTP
              address.
================================================================================
#>
#Requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DistributionListIdentity = '<DistributionListIdentity>'
)

# List all members of a specific distribution list
$Members = (Get-DistributionGroupMember $DistributionListIdentity).Identity
foreach ($Member in $Members) {
    Get-Mailbox $Member | Select-Object DisplayName, PrimarySMTPAddress
}