<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Remove-AzureADInactiveDevices
SYNOPSIS    : Remove all inactive AzureAD Devices Inactivity Threshold can be
              given
DESCRIPTION : Remove all inactive AzureAD Devices Inactivity Threshold can be
              given
================================================================================
#>
#Requires -Modules AzureAD
[CmdletBinding(ConfirmImpact = 'Low')]
param
(
   [Parameter(ValueFromPipelineByPropertyName = $true,
      ValueFromRemainingArguments = $true)]
   [ValidateNotNullOrEmpty()]
   [Alias('Threshold', 'Days')]
   [int]
   $TimeFrame = 90
)

begin
{
   # Build a Threshold value
   $InactivityTimeFrame = ((Get-Date).AddDays(-$TimeFrame))
}

process
{
   # Get all AzureAD Devices
   $paramGetAzureADDevice = @{
      All           = $true
      ErrorAction   = 'SilentlyContinue'
      WarningAction = 'SilentlyContinue'
   }
   $null = ((Get-AzureADDevice @paramGetAzureADDevice | Where-Object -FilterScript {
            $PSItem.ApproximateLastLogonTimeStamp -le $InactivityTimeFrame
         }) | ForEach-Object -Process {
         # Inform the operator
         Write-Output -InputObject ('Removing device {0}' -f $PSItem.ObjectId)

         try
         {
            # OPTIONAL: Disable the device (just to make it correct)
            $paramSetAzureADDevice = @{
               ObjectId       = $PSItem.ObjectId
               AccountEnabled = $false
               IsCompliant    = $false
               IsManaged      = $false
               Verbose        = $true
               ErrorAction    = 'SilentlyContinue'
               WarningAction  = 'SilentlyContinue'
            }
            $null = (Set-AzureADDevice @paramSetAzureADDevice)
         }
         catch
         {
            # AzureAD commands seems to ignore the ErrorAction parameter
            Write-Verbose -Message 'OK'
         }

         try
         {
            # Remove the device to get a clean device list
            $paramRemoveAzureADDevice = @{
               ObjectId      = $PSItem.ObjectId
               Verbose       = $true
               ErrorAction   = 'SilentlyContinue'
               WarningAction = 'SilentlyContinue'
            }
            $null = (Remove-AzureADDevice @paramRemoveAzureADDevice)
         }
         catch
         {
            # AzureAD commands seems to ignore the ErrorAction parameter
            Write-Verbose -Message 'OK'
         }
      })
}