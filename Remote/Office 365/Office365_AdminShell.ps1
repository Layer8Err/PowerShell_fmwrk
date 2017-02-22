#ExecuteOutOpen
######################################################################
# Connect to 365 online so that we can run
# commands in the 365 environment
# 
# You will still need to know what Office 365 commands to use...
######################################################################
## Install Azure Active Directory Module for Windows PowerShell first
# http://go.microsoft.com/fwlink/p/?linkid=236297
# https://www.microsoft.com/en-us/download/details.aspx?id=28177
# This must already be installed
######################################################################
$ConnectionUri = "https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS"
######################################################################
if ($365UserCredentials) {} else {
    Write-Host "Please enter your Office 365 e-mail address and password." -ForegroundColor White
    $365UserCredentials = Get-Credential
}
function Connect365 {
    Import-Module MSOnline
    ## Create 365 Session ##
    $365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $365UserCredentials -Authentication Basic -AllowRedirection
    Import-PSSession $365Session
    Connect-MSolService -Credential $365UserCredentials
    $365connected = $true
}

#Determine if connected in this tab
if ($365connected) {
    if ($365connected -eq $false){
        Connect365
        $365connected = $true
    }
} else {
    Connect365
    $365connected = $true
}
$users = Get-Mailbox -RecipientTypeDetails UserMailbox
$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox
