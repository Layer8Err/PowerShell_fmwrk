#Execute
#############################################
# This script is written to totally reset 
# Windows Networking Components
# It should be run as admin on the affected
# computer.
#
#############################################
Echo "Disabling Network Sharing..."

$MainAdapter = Get-NetAdapter | Where-Object {$_.MediaConnectionState -eq 'Connected' -and $_.PhysicalMediaType -ne 'Unspecified'} | Sort-Object LinkSpeed -Descending

$m = New-Object -ComObject HNetCfg.HNetShare

$c = $m.EnumEveryConnection |% { $m.NetConnectionProps.Invoke($_).Guid }
$co = $m.EnumEveryConnection |? { $m.NetConnectionProps.Invoke($_).Guid -eq $c }
$config = $m.INetSharingConfigurationForINetConnection.Invoke($co)

Write-Output $config.SharingEnabled

## EnableSharing (0 = public, 1 = private)
#$config.EnableSharing(1)
$config.DisableSharing()

Echo "Resetting winsock..."
netsh winsock reset
# reset winsock entries
Echo "Resetting winsock catalog entries..."
netsh winsock reset catalog
# Reset TCP/IP stack
Echo "Resetting TCP/IP stack..."
netsh int ip reset reset.log hit
Echo "Resetting IPv4..."
netsh int ipv4 reset

ECHO "Performing additional netsh stuff..."
netsh int tcp set heuristics disabled
netsh int tcp set global autotuninglevel=disable
netsh int tcp set global rss=enabled
netsh int tcp show global

# Flush DNS
Echo "Flushing DNS cache..."
ipconfig /flushdns
# Flush arp table
Echo "Flushing ARP table..."
arp -d *
# Renew IP
Echo "Renewing IP..."
ipconfig /release
ipconfig /renew

# Restart Adapter
Echo "Restarting network adapter..."
$adapters = Get-NetAdapter
Restart-NetAdapter -Name $adapters.Name

ECHO "Changing network status to Private..."
$networkListManager = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}'))
$connections = $networkListManager.GetNetworkConnections()
$connections | %{$_.GetNetwork().SetCategory(1)} 
Write-Host "You should reboot the PC now"
$reboot = Read-Host "Reboot now? [y/N]"
if ($reboot.ToLower().SubString(0,1) -eq "y") {
	Write-Host "Rebooting..." -ForegroundColor Yellow
	shutdown /t 030 /r /f /c "Rebooting your PC to clean the network interface"
	pause
}
