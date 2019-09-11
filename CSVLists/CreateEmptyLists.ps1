#Execute
###############################################################################################
# Create empty CSV files for lists of computers
# These files can be manually populated later
###############################################################################################

if (!(Test-Path $env:SERVERSLIST)){
    Write-Host "Creating dummy servers.csv file..." -ForegroundColor Yellow
    Write-Output "user,computer" > $env:SERVERSLIST
    Write-Output "user,server1" >> $env:SERVERSLIST
    Write-Output "user,server2" >> $env:SERVERSLIST
    Write-Output "user,server3" >> $env:SERVERSLIST
}
if (!(Test-Path $env:COMPUTERSLIST)){
    Write-Host "Creating dummy userPCs.csv file..." -ForegroundColor Yellow
    Write-Output "user,computer" > $env:COMPUTERSLIST
    Write-Output "user,computer1" >> $env:COMPUTERSLIST
    Write-Output "user,computer2" >> $env:COMPUTERSLIST
    Write-Output "user,computer3" >> $env:COMPUTERSLIST
}
