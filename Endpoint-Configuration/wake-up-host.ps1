<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : wake-up-host
SYNOPSIS    : This PowerShell script sends a magic UDP packet to a computer to
              wake him up (requires the target computer to have
              Wake-on-LAN activated).
DESCRIPTION : This PowerShell script sends a magic UDP packet to a computer to
              wake him up (requires the target computer to have
              Wake-on-LAN activated).
================================================================================
#>
param([string]$macAddr = "", [string]$ipAddr = "", [int]$udpPort = 9, [int]$numTimes = 3)
	
function Send-WOL { param([string]$macAddr, [string]$ipAddr, [int]$udpPort) 
	$broadcastAddr = [Net.IPAddress]::Parse($ipAddr) 
  
	$macAddr = (($macAddr.replace(":","")).replace("-","")).replace(".","") 
	$target = 0,2,4,6,8,10 | % {[convert]::ToByte($macAddr.substring($_,2),16)} 
	$packet = (,[byte]255 * 6) + ($target * 16) 
  
	$UDPclient = New-Object System.Net.Sockets.UdpClient 
	$UDPclient.Connect($broadcastAddr, $udpPort) 
	[void]$UDPclient.Send($packet, 102)  
} 

try {
	if ($macAddr -eq "" ) { $macAddr = Read-Host "Enter the host's MAC address, e.g. 11:22:33:44:55:66" }
	if ($ipAddr -eq "" ) { $ipAddr = Read-Host "Enter the subnet address, e.g. 192.168.178.255" }

	for ($i = 0; $i -lt $numTimes; $i++) {
		Send-WOL $macAddr.Trim() $ipAddr.Trim() $udpPort
		Start-Sleep -milliseconds 100	
	}
	"✅ Sent magic packet to $ipAddr, UDP port $($udpPort) ($($numTimes)x). The device is up in a minute."
	exit 0 # success
} catch {
	"⚠️ ERROR: $($Error[0]) (script line $($_.InvocationInfo.ScriptLineNumber))"
	exit 1
}