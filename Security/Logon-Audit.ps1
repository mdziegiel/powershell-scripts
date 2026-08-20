<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Logon-Audit
SYNOPSIS    : Audits Windows logon and logoff events from Security logs and can
              send the results to Microsoft Teams.
DESCRIPTION : Audits Windows logon and logoff events from Security logs and can
              send the results to Microsoft Teams.
================================================================================
#>
[CmdletBinding()]
Param(
    [alias("L")]
    $LogPathUsr,
    [Alias("Webhook")]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [string]$Webh,
    [switch]$Logon,
    [switch]$Logoff,
    [switch]$Help)

If ($PSBoundParameters.Values.Count -eq 0 -or $Help)
{
    Write-Host -Object "Usage:
    From a terminal run: [path\]Logon-Audit.ps1 -Logon -L [path]
    The above command will record a logon event for the currently logged on user to the log file and also to Teams.

    Use -Logoff to log a logoff event.
    
    To send events to a webhook:
    Specify a txt file containing the webhook URI with -Webhook [path\]webhook.txt
"
}

else {
    ## If logging is configured, set the log file name.
    If ($LogPathUsr)
    {
        ## Clean User entered string
        $LogPath = $LogPathUsr.trimend('\')

        ## Make sure the log directory exists.
        If ((Test-Path -Path $LogPath) -eq $False)
        {
            New-Item $LogPath -ItemType Directory -Force | Out-Null
        }

        $LogFile = "Logon-Audit.log"
        $Log = "$LogPath\$LogFile"
    }

    ## Function to get date in specific format.
    Function Get-DateFormat
    {
        Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    ## Function for logging.
    Function Write-Log($Type,$Evt)
    {
        If ($Type -eq "Logon")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [LOGON] $Evt"
            }
        }

        If ($Type -eq "Logoff")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [LOGOFF] $Evt"
            }
        }

        If ($Type -eq "Err")
        {
            If ($LogPathUsr)
            {
                Add-Content -Path $Log -Encoding ASCII -Value "$(Get-DateFormat) [ERROR] $Evt"
            }

            Write-Host -ForegroundColor Red -BackgroundColor Black -Object "$(Get-DateFormat) [ERROR] $Evt"
        }
    }

    If ($Logon -eq $false -And $Logoff -eq $false)
    {
        Write-Log -Type Err -Evt "Not Configured to do anything. Specify -Logon or -Logoff."
        Exit
    }

    # If the -logon switch is used, register it as a logon.
    If ($Logon)
    {
        Write-Log -Type Logon -Evt "Device: $env:COMPUTERNAME, Domain: $env:userdomain, Username: $env:username"

        If ($Webh)
        {
            $EStatus = "Logon"
        }
    }

    # If the -logon switch is used, register it as a logoff.
    If ($Logoff)
    {
        Write-Log -Type Logoff -Evt "Device: $env:COMPUTERNAME, Domain: $env:userdomain, Username: $env:username"

        If ($Webh)
        {
            $EStatus = "Logoff"
        }
    }

    # If the teams switch is used, get the webhook uri from the txt file.
    If ($Webh)
    {
        $uri = Get-Content $Webh

        # Create an array for the results.
        $ResultArr = @()

        $ResultArr += New-Object PSObject -Property @{
            facts = @(
                @{
                    name = 'User:'
                    value = $env:username
                },
                @{
                    name = 'Event:'
                    value = $EStatus
                },
                @{
                    name = 'Device:'
                    value = $env:COMPUTERNAME
                },
                @{
                    name = 'Domain:'
                    value = $env:userdomain
                }
            )
        }

        # If the result is not empty, put array together for sending to teams.
        If ($Null -ne $ResultArr)
        {
            $Body = ConvertTo-Json -Depth 8 @{
            text  = "An event occurred."
            sections = $ResultArr
            title = "Logon Audit Utility"
            }

            Invoke-RestMethod -Uri $Uri -Method Post -body $Body -ContentType 'application/json'
        }
    }
}

## End