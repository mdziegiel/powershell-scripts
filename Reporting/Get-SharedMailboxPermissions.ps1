<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get-SharedMailboxPermissions
SYNOPSIS    : Connects to Exchange Online and enumerates shared mailbox
              permissions.
DESCRIPTION : Connects to Exchange Online and enumerates shared mailbox
              permissions.
================================================================================
#>
[CmdletBinding()]
param(
    [PSCredential]$Credential
)

Import-Module ExchangeOnlineManagement -ErrorAction Stop

if ($Credential) {
    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
} else {
    Connect-ExchangeOnline -ShowBanner:$false
}

# Get all shared mailboxes
$sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Initialize an array to store mailbox permissions
$mailboxPermissions = @()

# Function to extract creation date from mailbox name
function Extract-CreationDate {
    param (
        [string]$identity
    )
    if ($identity -match '\d{8}\d{6}$') {
        return [datetime]::ParseExact($identity.Substring($identity.Length - 14), "yyyyMMddHHmmss", $null)
    }
    else {
        return $null
    }
}

# Loop through each shared mailbox
foreach ($mailbox in $sharedMailboxes) {
    try {
        # Get mailbox permissions for the current mailbox
        $permissions = Get-MailboxPermission -Identity $mailbox.Identity
        
        # Extract creation date from mailbox name
        $creationDate = Extract-CreationDate -identity $mailbox.Identity
        
        # Filter and select the necessary properties
        $filteredPermissions = $permissions | Select-Object @{Name='Mailbox';Expression={$mailbox.DisplayName}}, User, AccessRights, @{Name='CreatedOn';Expression={$creationDate}} | Where-Object { $_.User -like '*@*' }
        
        # Add the filtered permissions to the array
        $mailboxPermissions += $filteredPermissions
    }
    catch [Microsoft.Exchange.Configuration.Tasks.ManagementObjectAmbiguousException] {
        Write-Warning "Ambiguous recipient for mailbox: $($mailbox.Identity) - $_"
    }
    catch {
        Write-Warning "Could not retrieve permissions for mailbox: $($mailbox.Identity) - $_"
    }
}

# Export the collected permissions to a CSV file
$mailboxPermissions | Export-Csv -Path (Join-Path $env:TEMP 'SharedMailboxes.csv') -NoTypeInformation