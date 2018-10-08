#Execute
###############################################################################################
# Clear Outlook cache
# 
# Clear Outlook cache on remote PC
# This solves some Outlook issues
###############################################################################################
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName

###############################################################################################
#######################################*Begin Script*##########################################
###############################################################################################

$user = @()
$computer = @()
$pcname = Read-Host 'PC-Name to clear Outlook cache'

if ($cred) {} else { $cred = Get-Credential $adminName }

#import list of usernames and computernames from .csv
#save into two arrays
Import-Csv $pcList | ForEach-Object {
        $user += $_.user
        $computer += $_.computer
}

#loop through arrays to find username
for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i]
    if ($icomputer -eq $pcname) { $userName = $iuser }
}

$script ='Start-Sleep -s 1 ;
    Kill -name *skype* -Force ;
    Kill -Name *lync* -Force ;
    Kill -Name *UCA* -Force ;
    Kill -Name *UcMapi* -Force ;
    Kill -Name *Outlook* -Force ;
    Start-Sleep -s 2 ;
    Kill -Name *OUTLOOK* -Force ;
    Start-Sleep -s 3 ;
    $tempPath = "C:\Users\$userName\AppData\Local\Microsoft\Outlook\*.ost" ;
    Remove-Item $tempPath -Force ;
    msg ' + $userName + '"Please re-open Skype for Business and Outlook"'

#Execute Outlook Cache Clear
$scriptBlock = [ScriptBlock]::Create($script)
Invoke-Command -Computer $icomputer -Credential $cred -ScriptBlock $scriptBlock
