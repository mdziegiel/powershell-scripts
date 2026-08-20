<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Get-Office365Endpoints
SYNOPSIS    : Microsoft updates the Office 365 IP address and FQDN entries at
              the end of each month and occasionally out of the
              cycle for operational or support requirements
DESCRIPTION : Microsoft updates the Office 365 IP address and FQDN entries at
              the end of each month and occasionally out of the
              cycle for operational or support requirements. This
              function uses the new JSON based Webserice instead
              of…
================================================================================
#>
function Get-Office365Endpoints
{
   [CmdletBinding(ConfirmImpact = 'None')]
   [OutputType([psobject])]
   param
   (
      [Parameter(ValueFromPipeline,
         ValueFromPipelineByPropertyName)]
      [ValidateSet('Worldwide', 'USGovDoD', 'USGovGCCHigh', 'China', 'Germany', IgnoreCase = $true)]
      [ValidateNotNullOrEmpty()]
      [string]
      $Instance = 'Worldwide',
      [Parameter(ValueFromPipeline)]
      [ValidateSet('All', 'Common', 'Exchange', 'SharePoint', 'Skype', IgnoreCase = $true)]
      [ValidateNotNullOrEmpty()]
      [Alias('ServiceAreas')]
      [string]
      $Services = 'All',
      [Parameter(ValueFromPipeline)]
      [Alias('TenantName')]
      [string]
      $Tenant = $null,
      [Parameter(ValueFromPipeline)]
      [switch]
      $NoIPv6,
      [Parameter(ValueFromPipeline)]
      [switch]
      $ExpressRoute,
      [ValidateSet('All', 'Optimize', 'Allow', 'Default', IgnoreCase = $true)]
      [string[]]
      $Category,
      [Parameter(ValueFromPipeline)]
      [switch]
      $Required,
      [Parameter(ValueFromPipeline)]
      [ValidateSet('All', 'IPv4', 'IPv6', 'URLs', IgnoreCase = $true)]
      [string]
      $Output = 'All',
      [Parameter(ValueFromPipeline,
         ValueFromPipelineByPropertyName)]
      [Alias('ForceDownload')]
      [switch]
      $SkipVersionCheck = $false
   )

   begin
   {
      #region MakeIPv6Plausible
      if (($NoIPv6) -and ($Output -eq 'IPv6'))
      {
         # This makes no sense, and we totally ignore to do it!
         Write-Error -Message 'The selected parameters make no sense; we cannot continue!' -ErrorAction Stop

         # We should never reach this point!
         break
      }
      #endregion MakeIPv6Plausible

      #region CategoryTweaker
      if ((! $Category) -or ($Category -eq 'All'))
      {
         Write-Verbose -Message 'We get all categories.'

         # Set to all
         $Category += 'Optimize', 'Allow', 'Default'
      }
      #endregion CategoryTweaker

      #region TweakOutputHandler

      <#
         TODO: Make a simpler solution for that!
      #>
      switch ($Output)
      {
         'All'
         {
            Write-Verbose -Message 'Dump all Infos (IPv4, IPv6, and URLs)'

            $outIPv4 = $true
            $outIPv6 = $true
            $outURLs = $true
         }
         'IPv4'
         {
            Write-Verbose -Message 'Dump IPv4 Infos'

            $outIPv4 = $true
            $outIPv6 = $false
            $outURLs = $false
         }
         'IPv6'
         {
            Write-Verbose -Message 'Dump IPv6 Infos'

            $outIPv4 = $false
            $outIPv6 = $true
            $outURLs = $false
         }
         'URLs'
         {
            Write-Verbose -Message 'Dump URLs Infos'

            $outIPv4 = $false
            $outIPv6 = $false
            $outURLs = $true
         }
      }
      #endregion TweakOutputHandler

      #region ConfigurationVariables
      # Web service root URL
      $BaseURI = 'https://endpoints.office.com'
      Write-Verbose -Message ('We use {0} as Base URL' -f $BaseURI)

      # Path where client ID and latest version number will be stored
      <#
        TODO: Move the Location to a parameter
    #>
      $datapath = $Env:TEMP + '\O365_endpoints_' + $Instance + '_latestversion.txt'

      Write-Verbose -Message ('We save the Endpoint Version Information to {0}' -f $datapath)
      #endregion ConfigurationVariables

      #region LocalVersionChecker
      # fetch client ID and version if data file exists; otherwise create new file
      if (Test-Path -Path $datapath)
      {
         Write-Verbose -Message 'We get the information from Microsoft...'

         # Read the File
         $content = (Get-Content -Path $datapath)

         # Get the Info
         $clientRequestId = $content[0]
         $lastVersion = $content[1]

         # Cleanup
         $content = $null
      }
      else
      {
         Write-Verbose -Message 'Old version information file exists, start to gather the Info!'

         # Create a GUID
         $clientRequestId = [GUID]::NewGuid().Guid

         # Dummy Data
         $lastVersion = '0000000000'

         # Save the local info
         try
         {
            @($clientRequestId, $lastVersion) | Out-File -FilePath $datapath -ErrorAction Stop
         }
         catch
         {
            # Write the complete error if we have verbose turned on
            Write-Verbose -Message $_

            # Our Error test
            Write-Error -Message ('Unable to write Datafile: {0}' -f $datapath) -ErrorAction Stop

            # We should never reach this point!
            break
         }
      }
      #endregion LocalVersionChecker

      #region RemoteVersionChecker
      # Call version method to check the latest version, and pull new data if version number is different
      try
      {
         # Splat the parameters
         $GetVersionParams = @{
            Uri           = ($BaseURI + '/version/' + $Instance + '?clientRequestId=' + $clientRequestId)
            Method        = 'Get'
            ErrorAction   = 'Stop'
            WarningAction = 'SilentlyContinue'
         }

         Write-Verbose -Message ('We use {0} as request URI.' -f ($GetVersionParams.Uri))

         $version = (Invoke-RestMethod @GetVersionParams)
      }
      catch
      {
         # Write the complete error if we have verbose turned on
         Write-Verbose -Message $_

         # Our Error test
         Write-Error -Message 'Unable to get the new Office 365 Endpoint Information' -ErrorAction Stop

         # We should never reach this point!
         break
      }
      #endregion RemoteVersionChecker
   }

   process
   {
      #region VersionCompare
      if (($SkipVersionCheck -eq $true) -or ($version.latest -gt $lastVersion))
      {
         Write-Verbose -Message ('New version of Office 365 {0} endpoints detected' -f $Instance)

         # Write the new version number to the data file
         try
         {
            @($clientRequestId, $version.latest) | Out-File -FilePath $datapath -ErrorAction Stop
         }
         catch
         {
            # Write the complete error if we have verbose turned on
            Write-Verbose -Message $_

            # Our Error test
            Write-Error -Message ('Unable to write Datafile: {0}' -f $datapath) -ErrorAction Stop

            # We should never reach this point!
            break
         }
         #endregion VersionCompare

         #region GetTheEndpoints
         try
         {
            # Set the default URI
            $requestURI = ($BaseURI + '/endpoints/' + $Instance + '?clientRequestId=' + $clientRequestId)

            switch ($Services)
            {
               'All'
               {
                  # We get all
               }
               'Common'
               {
                  # Append to the URI
                  $requestURI = ($requestURI + '&ServiceAreas=Common')
               }
               'Exchange'
               {
                  # Append to the URI
                  $requestURI = ($requestURI + '&ServiceAreas=Exchange')
               }
               'SharePoint'
               {
                  # Append to the URI
                  $requestURI = ($requestURI + '&ServiceAreas=SharePoint')
               }
               'Skype'
               {
                  # Append to the URI
                  $requestURI = ($requestURI + '&ServiceAreas=Skype')
               }
            }

            if ($Tenant)
            {
               # Append to the URI - Build URL for the Tenant
               $requestURI = ($requestURI + '&TenantName=' + $Tenant)
            }

            if ($NoIPv6)
            {
               # Append to the URI - Exclude IPv6 addresses from the output
               $requestURI = ($requestURI + '&NoIPv6')

               Write-Verbose -Message 'IPv6 addresses are excluded from the output! IPv6 is the future, think about an adoption soon.'
            }

            # Do our job and get the data via Rest Request
            Write-Verbose -Message ('We request the following URI: {0}' -f $requestURI)

            $endpointSetsParams = @{
               Uri           = $requestURI
               Method        = 'Get'
               ErrorAction   = 'Stop'
               WarningAction = 'SilentlyContinue'
            }
            $endpointSets = (Invoke-RestMethod @endpointSetsParams)
         }
         catch
         {
            # Write the complete error if we have verbose turned on
            Write-Verbose -Message $_

            # Our Error test
            Write-Error -Message 'Unable to get the new Office 365 Endpoint Information' -ErrorAction Stop

            # We should never reach this point!
            break
         }
         #endregion GetTheEndpoints

         #region FilterURLs

         if ($outURLs)
         {
            $flatUrls = $endpointSets | ForEach-Object -Process {
               $endpointSet = $PSItem
               $urls = $(if ($endpointSet.urls.Count -gt 0)
                  {
                     $endpointSet.urls
                  }
                  else
                  {
                     @()
                  }
               )

               # Cleanup
               $urlCustomObjects = @()

               if ($endpointSet.category -in ($Category))
               {
                  $urlCustomObjects = $urls | ForEach-Object -Process {
                     # Ordered is slower, but we like it this way
                     [PSCustomObject][ordered]@{
                        id           = $endpointSet.id
                        serviceArea  = $endpointSet.serviceArea
                        DisplayName  = $endpointSet.serviceAreaDisplayName
                        url          = $PSItem
                        tcpPorts     = $endpointSet.tcpPorts
                        udpPorts     = $endpointSet.udpPorts
                        expressRoute = $endpointSet.expressRoute
                        category     = $endpointSet.category
                        required     = $endpointSet.required
                        notes        = $endpointSet.notes
                     }
                  }
               }

               # Only ExpressRoute enabled Objects?
               if ($ExpressRoute)
               {
                  $urlCustomObjects = $urlCustomObjects | Where-Object -FilterScript {
                     $urlCustomObjects.expressRoute -eq $true
                  }
               }

               # Only required to have connectivity for Office 365 to be supported
               if ($Required)
               {
                  $urlCustomObjects = $urlCustomObjects | Where-Object -FilterScript {
                     $urlCustomObjects.required -eq $true
                  }
               }

               # Dump
               $urlCustomObjects
            }
         }
         #endregion FilterURLs

         #region FilterIPv4
         if ($outIPv4)
         {
            $flatIpv4 = $endpointSets | ForEach-Object -Process {
               $endpointSet = $PSItem
               $ips = $(if ($endpointSet.ips.Count -gt 0)
                  {
                     $endpointSet.ips
                  }
                  else
                  {
                     @()
                  }
               )

               # IPv4 strings have dots while IPv6 strings have colons
               $IPv4 = $ips | Where-Object -FilterScript {
                  $PSItem -like '*.*'
               }

               # Cleanup
               $ipCustomObjects = @()

               if ($endpointSet.category -in ($Category))
               {
                  $ipCustomObjects = $IPv4 | ForEach-Object -Process {
                     # Ordered is slower, but we like it this way
                     [PSCustomObject][ordered]@{
                        id           = $endpointSet.id
                        serviceArea  = $endpointSet.serviceArea
                        DisplayName  = $endpointSet.serviceAreaDisplayName
                        ip           = $PSItem
                        tcpPorts     = $endpointSet.tcpPorts
                        udpPorts     = $endpointSet.udpPorts
                        expressRoute = $endpointSet.expressRoute
                        category     = $endpointSet.category
                        required     = $endpointSet.required
                        notes        = $endpointSet.notes
                     }
                  }
               }

               # Dump
               $ipCustomObjects
            }
         }
         #endregion FilterIPv4

         #region FilterIPv6
         if ($outIPv6)
         {
            $flatIpv6 = $endpointSets | ForEach-Object -Process {
               $endpointSet = $PSItem
               $ips = $(if ($endpointSet.ips.Count -gt 0)
                  {
                     $endpointSet.ips
                  }
                  else
                  {
                     @()
                  }
               )

               # IPv4 strings have dots while IPv6 strings have colons
               $IPv6 = $ips | Where-Object -FilterScript {
                  $PSItem -like '*:*'
               }

               # Cleanup
               $ipCustomObjects = @()

               if ($endpointSet.category -in ($Category))
               {
                  $ipCustomObjects = $IPv6 | ForEach-Object -Process {
                     # Ordered is slower, but we like it this way
                     [PSCustomObject][ordered]@{
                        id           = $endpointSet.id
                        serviceArea  = $endpointSet.serviceArea
                        DisplayName  = $endpointSet.serviceAreaDisplayName
                        ip           = $PSItem
                        tcpPorts     = $endpointSet.tcpPorts
                        udpPorts     = $endpointSet.udpPorts
                        expressRoute = $endpointSet.expressRoute
                        category     = $endpointSet.category
                        required     = $endpointSet.required
                        notes        = $endpointSet.notes
                     }
                  }
               }

               # Dump
               $ipCustomObjects
            }
         }
         #endregion FilterIPv4
      }
   }

   end
   {
      if (($SkipVersionCheck -eq $true) -or ($version.latest -gt $lastVersion))
      {
         #region DumpIPv4
         if ($outIPv4)
         {
            Write-Verbose -Message 'Office 365 IPv4 IP Address Ranges'

            ($flatIpv4 | Sort-Object -Property id)
         }
         #endregion DumpIPv4

         #region DumpIPv6
         if ($outIPv6)
         {
            Write-Verbose -Message 'Office 365 IPv6 IP Address Ranges'

            ($flatIpv6 | Sort-Object -Property id)
         }
         #endregion DumpIPv6

         #region DumpURLs
         if ($outURLs)
         {
            Write-Verbose -Message 'Office 365 URLs'

            ($flatUrls | Sort-Object -Property id)
         }
         #endregion DumpURLs
      }
      else
      {
         #region DumpInfoNothing
         <#
          This 'else' loop is here as a placeholder in this script!
          We use this in the commercial version (function within the commercial module)
      #>

         Write-Output -InputObject ('The {0} Office 365 endpoints are up-to-date' -f $Instance)
         #endregion DumpInfoNothing
      }
   }

   #region CHANGELOG
   <#
      CHANGELOG:
      0.8.6 - 2019.01-04:
      [CHANGE] Converted back to a function to make it easier for me (no more need to adopt between my sources)
      [ADD] Start to add regions

      0.8.5 - 2018-10-04:
      [FIX] Fix the Output to reflect the correct Instance name (PSMO365-48)
      [ADD] Add -SkipVersionCheck Switch to force the download. As request by @mikes-gh in #4 in GitHub (PSMO365-49)
      [FIX] Fix a view typos and errors

      0.8.4 - 2018-08-29:
      [ADD] Exchange Online Example added (Source http://www.powershell.no/exchange/online,office/365,powershell/2018/08/26/automate-office365-ip-address-handling.html)
      [CHANGE] Tweaks (after internal code review and refactoring)

      0.8.3 - 2018-08-20 - Unreleased:
      [ADD] We added a few more verbose outputs. Verbose Implementation us based upon request. (PSMO365-43)
      [CHANGE] Region name change

      0.8.2 - 2018-08-19:
      [ADD] Regions added to make the code more readable within code editors (PSMO365-47)
      [FIX] A few typos in the descriptions where fixed - No change to any code or logic

      0.8.1 - 2018-08-19:
      [FIX] Add missing OutputType (PSMO365-41)
      [CHANGE] datafile name tweaked (PSMO365-42)
      [ADD] Missing NoIPv6 switch function implemented (PSMO365-44)
      [ADD] New Example for NoIPv6 switch (PSMO365-45)
      [ADD] A few more links
      [ADD] Info about the datafile (PSMO365-42)
      [ADD] Embed a few things as comment - Due to the separation from the Module
      [ADD] This changelog within the code - Reflect the changes within the dedicated function (PSMO365-46)

      0.8.0 - 2018-08-18:
      [INIT] Initial public release
  #>
   #endregion CHANGELOG
}

#region LICENSE
<#
   BSD 3-Clause License

   Copyright (c) 2022, enabling Technology
   All rights reserved.

   Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

   1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
   2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
   3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>
#endregion LICENSE

#region DISCLAIMER
<#
   DISCLAIMER:
   - Use at your own risk, etc.
   - This is open-source software, if you find an issue try to fix it yourself. There is no support and/or warranty in any kind
   - This is a third-party Software
   - The developer of this Software is NOT sponsored by or affiliated with Microsoft Corp (MSFT) or any of its subsidiaries in any way
   - The Software is not supported by Microsoft Corp (MSFT)
   - By using the Software, you agree to the License, Terms, and any Conditions declared and described above
   - If you disagree with any of the terms, and any conditions declared: Just delete it and build your own solution
#>
#endregion DISCLAIMER