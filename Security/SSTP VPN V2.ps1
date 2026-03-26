<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Configure SSTP VPN Connection
SYNOPSIS    : Creates or updates an all-user VPN connection
DESCRIPTION : Creates or updates a Windows VPN connection using the
              specified connection settings, applies DNS suffix
              configuration, enables split tunneling, and adds
              static routes for internal network access

ORGANIZATION: Hanskissle
==============================================================================
#>

# If PowerShell supports VPN configuration, apply VPN configuration
if (Get-Command 'Get-VpnConnection') {
    # If VPN exists, update VPN settings
    if (Get-VpnConnection -Name $Name -AllUserConnection -ErrorAction SilentlyContinue) {
        Set-VpnConnection -Name $Name -DnsSuffix $DnsSuffix -AllUserConnection
        add-vpnconnectionroute -connectionname $Name -destinationprefix 172.17.1.0/24 -passthru
        add-vpnconnectionroute -connectionname $Name -destinationprefix 172.16.1.0/24 -passthru
        #Write-Host "VPN already exists"
    }
    # Else, create VPN connection
    else {
       # Add-VpnConnection -name "Hanskissle VPN" -serveraddress "vpn.hanskissle.com" -TunnelType SSTP -splittunneling -remembercredential -UseWinlogonCredential -alluserconnection -encryptionlevel "required"
        Add-VpnConnection -Name $Name -AllUserConnection -ServerAddress $ServerAddress -TunnelType $TunnelType -EncryptionLevel $EncryptionLevel -DnsSuffix $DnsSuffix -SplitTunneling -UseWinlogonCredential -RememberCredential -force
        add-vpnconnectionroute -connectionname $Name -destinationprefix 172.17.1.0/24 -passthru
        add-vpnconnectionroute -connectionname $Name -destinationprefix 172.16.1.0/24 -passthru
    }
    return Get-VpnConnection -Name $Name -AllUserConnection
    exit
}
# Else, exit with failure code
else {
  	return "Client does not support VpnClient cmdlets"
	exit 1
}

