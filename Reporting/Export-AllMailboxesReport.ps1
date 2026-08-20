<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Export-AllMailboxesReport
SYNOPSIS    : Retrieves mailbox details from Exchange Online, including aliases,
              mailbox size, full access permissions, send-as
              permissions, and forwarding configuration, then
              exports the results to CSV for reporting and audit…
DESCRIPTION : Retrieves mailbox details from Exchange Online, including aliases,
              mailbox size, full access permissions, send-as
              permissions, and forwarding configuration, then
              exports the results to CSV for reporting and audit…
================================================================================
#>
#Requires -Modules ExchangeOnlineManagement
$excludedEmails = @(
    "excluded1@contoso.com",
    "excluded2@contoso.com"
)

Get-Mailbox -ResultSize Unlimited |
Where-Object { $excludedEmails -notcontains $_.PrimarySmtpAddress.ToString().ToLower() } |
ForEach-Object {
    $mb = $_
    $guidString = $mb.Guid.Guid

    # Primary stats
    $mbStats = Get-MailboxStatistics -Identity $guidString -ErrorAction SilentlyContinue
    $mailboxSize = if ($mbStats) { $mbStats.TotalItemSize.ToString() } else { "N/A" }

    # Filter aliases: exclude primary, onmicrosoft, sip, x500, exch.* addresses
    $filteredAliases = $mb.EmailAddresses |
        Where-Object {
            ($_.ToString().ToLower().StartsWith("smtp:")) -and
            ($_.ToString().Split(":")[1].ToLower() -ne $mb.PrimarySmtpAddress.ToString().ToLower()) -and
            ($_.ToString().ToLower() -notlike "*sip:*") -and
            ($_.ToString().ToLower() -notlike "*x500:*") -and
            ($_.ToString().Split(":")[1].ToLower() -notlike "*@exch.*.com") -and
            ($_.ToString().Split(":")[1].ToLower() -notlike "*.onmicrosoft.com")
        } |
        ForEach-Object { $_.ToString().Split(":")[1] }

    # Extract tenant routing address
    $tenantRoutingAddress = $mb.EmailAddresses |
        Where-Object {
            ($_.ToString().ToLower().StartsWith("smtp:")) -and
            ($_.ToString().ToLower() -like "*@*.mail.onmicrosoft.com")
        } |
        ForEach-Object { $_.ToString().Split(":")[1] } |
        Select-Object -First 1

    # Full Access permissions
    $fullAccess = Get-MailboxPermission -Identity $guidString |
        Where-Object {
            $_.AccessRights -contains "FullAccess" -and
            -not $_.IsInherited -and
            $_.User -ne "NT AUTHORITY\SELF"
        } |
        Select-Object -ExpandProperty User

    $fullAccessNames = $fullAccess | ForEach-Object {
        $user = Get-User -Identity $_ -ErrorAction SilentlyContinue
        if ($user.DisplayName -and $user.DisplayName -notmatch "^quest") {
            $user.DisplayName
        }
    }

    # Send As permissions
    $sendAs = Get-RecipientPermission -Identity $guidString |
        Where-Object {
            $_.AccessRights -contains "SendAs" -and
            -not $_.IsInherited -and
            $_.Trustee -ne "NT AUTHORITY\SELF"
        } |
        Select-Object -ExpandProperty Trustee

    $sendAsNames = $sendAs | ForEach-Object {
        Get-User -Identity $_ -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty DisplayName
    }

    # Mailbox forwarding
    $forwardingAddress = if ($mb.ForwardingAddress) {
        try {
            (Get-Recipient -Identity $mb.ForwardingAddress -ErrorAction Stop).DisplayName
        } catch {
            $mb.ForwardingAddress.ToString()
        }
    } else { "" }

    $forwardingSmtp = if ($mb.ForwardingSmtpAddress) {
        $mb.ForwardingSmtpAddress.ToString()
    } else { "" }

    $deliverToMailboxAndForward = $mb.DeliverToMailboxAndForward

    [PSCustomObject]@{
        DisplayName                 = $mb.DisplayName
        PrimarySmtpAddress          = $mb.PrimarySmtpAddress
        Alias                       = ($filteredAliases -join "; ")
        TenantRoutingAddress        = $tenantRoutingAddress
        MailboxSize                 = $mailboxSize
        FullAccessUsers             = ($fullAccessNames -join "; ")
        SendAsUsers                 = ($sendAsNames -join "; ")
        ForwardingAddress           = $forwardingAddress
        ForwardingSmtpAddress       = $forwardingSmtp
        DeliverToMailboxAndForward  = $deliverToMailboxAndForward
    }
} | Export-Csv -Path .\AllMailboxesReport.csv -NoTypeInformation