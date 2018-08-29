#Execute
###############################################################################################
# Create empty CSV files for lists of computers
# These files can be manually populated later
###############################################################################################

$listsRoot = $env:PSFmwrkRoot + '\Remote\PCLists'

Set-Location $PSScriptRoot
$localDir = $pwd.Path
$settingsXMLFile = $localDir + '\'
if ($listsRoot.Length -gt 16){
    $localDir = $listsRoot
}

$serverlist = $localDir + '\servers.csv'
$userpclist = $localDir + '\userPCs.csv'

if (!(Test-Path $serverlist)){
    Write-Host $serverlist
    Write-Output "user,computer" > $serverlist
    Write-Output "user,server1" >> $serverlist
    Write-Output "user,server2" >> $serverlist
    Write-Output "user,server3" >> $serverlist
}
if (!(Test-Path $userpclist)){
    Write-Host $userpclist
    Write-Output "user,computer" > $userpclist
    Write-Output "user,computer1" >> $userpclist
    Write-Output "user,computer2" >> $userpclist
    Write-Output "user,computer3" >> $userpclist
}
