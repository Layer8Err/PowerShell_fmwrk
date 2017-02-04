#Execute
###############################################################################################
# Send Wake on LAN packet
#
# By default the WOL will be sent to broadcast
# this may not work in some cases if an IP is not defined
###############################################################################################
# This script was created by Barry Chum
# https://gallery.technet.microsoft.com/scriptcenter/Send-WOL-packet-using-0638be7b
# borrowed under MICROSOFT LIMITED PUBLIC LICENSE version 1.1

function Send-WOL {
    <# 
      .SYNOPSIS  
        Send a WOL packet to a broadcast address
      .PARAMETER mac
       The MAC address of the device that need to wake up
      .PARAMETER ip
       The IP address where the WOL packet will be sent to (default broadcast)
      .EXAMPLE 
       Send-WOL -mac 00:11:32:21:2D:11 -ip 192.168.8.255 
    #>

    param(
    [string]$mac,
    [string]$ip="255.255.255.255",
    [int]$port=7,
    [int]$Packets=2
    )
    #$ip = "255.255.255.255"
    $broadcast = [Net.IPAddress]::Parse($ip)
    #$macOrig = $mac
    $mac=(($mac.replace(":","")).replace("-","")).replace(".","")
    $target=0,2,4,6,8,10 | % {[convert]::ToByte($mac.substring($_,2),16)}
    $packet = (,[byte]255 * 6) + ($target * 16)
 
    $UDPclient = new-Object System.Net.Sockets.UdpClient
    $UDPclient.Connect($broadcast,$port)
    [void]$UDPclient.Send($packet, 102)
}

$pcMAC = Read-Host 'MAC of PC to wake (aa:bb:cc:dd:ee:ff)'
Send-WOL -mac $pcMac


