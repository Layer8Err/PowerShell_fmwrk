#ExecuteOut
###############################################################################################
# Connect to remote PC via HTTPS
#
# This script is intended to be used with PCs not connected to the domain
# The remote PC must already have an SSL certificate set up
# and must have WinRM enabled.
###############################################################################################

$pcname = Read-Host 'PC-Name to connect to (non-Domain)'
$user = Read-Host 'Username'
#$WANip = Read-Host 'WAN IP'
$port = Read-Host 'Port [5986])'
if ($port -eq ""){
    $port = "5986"
}

$user = $pcname + "\" + $user
$cred = Get-Credential $user
$sess = New-PSSession -ComputerName $pcname -Credential $cred -Port $port -UseSSL -AllowRedirection -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck)
Enter-PSSession $sess