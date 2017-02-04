#Execute
###############################################################################################
# Poll who is logged into a machine using jobs
#
# This Powershell script has been tested with Powershell 3.0
# Query all computers in CSV list for who is logged on.
###############################################################################################
$pcList = $env:COMPUTERSLIST # Path to list of computer-names and user-names
$adminName = $env:adminName  # Domain admin username (e.g. Domain\admin)
#######################################*Begin Script*##########################################
#Import-Module ActiveDirectory # Not needed since we use the log
if ($cred) {} else { $cred = Get-Credential $env:adminName }
$user = @()
$computer = @()
$computers = @()
$offlinePCs = @()

#import list of usernames and computernames from .csv
$userList = Import-CSV $pcList
$userList | ForEach-Object { $user += $_.user ; $computer += $_.computer }

function pollLoggedIn {
    [CmdletBinding()] Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [String]$pcname,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$uname
    )
    $getLoggedInBlockFunction = [ScriptBlock]::Create({
        function getLoggedIn {
            [CmdletBinding()] Param(
            [Parameter(Position = 0, Mandatory = $True)]
            [String]$Computer,
            [Parameter(Position = 1, Mandatory = $True)]
            [String]$User
            )
            $currentUsers = @()
            $currentUsers = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object Username).Username
            if ($currentUsers.Length -eq 0){
                $currentUsers = (Get-WmiObject Win32_Process | Where-Object Name -Match explorer -ErrorAction Stop).GetOwner().user # attempt to get logged in user from explorer process
            }
            $boottime = (Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object lastbootuptime).lastbootuptime
            $lastLogin = (Get-WinEvent -MaxEvents 1 -FilterHashtable @{logname='Security'; ProviderName='Microsoft-Windows-Security-Auditing'; Id='4624'; Data="$User"}).TimeCreated
            $properties = @{
                    'Computer'=$Computer;
                    'User'=$User;
                    'currentUsers'=$currentUsers;
                    'lastLogin'=$lastLogin;
                    'lastBoot'=$boottime}
            $computerstats = New-Object PSObject -Property $properties
            Return $computerstats
        }
    })
    $getLoggedInBlock = [ScriptBlock]::Create($getLoggedInBlockFunction.ToString() + "getLoggedIn -Computer $pcname -User $uname")
    #Write-Output $getLoggedInBlock.ToString() # if you want to see the scriptBlock
    Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $getLoggedInBlock -AsJob
}
function userLookup ([string]$pcname){ return $userList.user[[array]::IndexOf($userList.computer, $pcname)] }
for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i]
	$currentOpp = "Requesting logged in users on $($icomputer) ( $($iuser)`'s PC)"
	Write-Output $currentOpp
    if (!(Test-Connection -ComputerName $icomputer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { 
        $offlinePC = New-Object PSObject -Property @{'Offline'=$icomputer}
        $offlinePCs += $offlinePC
    } else {
        if ($icomputer -ne $env:COMPUTERNAME){
            pollLoggedIn -pcname $icomputer -uname $iuser
        }
    }
}
Start-Sleep -Seconds 5 # Wait for jobs to finish
$loop = $true
$count = 0
$njobs = (Get-Job -State Running).count
Clear-Host
Write-Host -NoNewLine "Waiting for all jobs ($njobs) to complete" -ForegroundColor White
While ($loop){
    if ((Get-Job -State Running).count -eq 0){ $loop = $false }
    if ($count -eq 50){ $loop = $false }
    $jobs = (Get-Job -State Running).count
    if ($jobs -ne $njobs){
        Clear-Host
        Write-Host -NoNewLine "Waiting for all jobs ($jobs) to complete" -ForegroundColor White
        for ($i = 0 ; $i -lt $count + 1 ; $i ++){ Write-Host -NoNewline "." -ForegroundColor White }
        $njobs = $jobs
    } else {
        Write-Host -NoNewLine "." -ForegroundColor White
    }
    Start-Sleep -Seconds 1 # Wait for any running jobs to finish
    $count = $count + 1
}
Clear-Host
Get-Job -State Completed | ForEach-Object {
    $computers += (Get-Job -Id $_.Id | Receive-Job -ErrorAction SilentlyContinue)
}
if (((Get-Job -State Failed).count -gt 0) -or ($offline.Length -gt 0) -or ((Get-Job -State Running).count -gt 0) -or ($offlinePCs.Offline.Count -gt 0)){
    if ((Get-Job -State Running).count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "PCs with unfinnished business:" -ForegroundColor Yellow
        (Get-Job -State Running).Location | ForEach-Object {
            $u = userLookup -pcname $_
            $p = $_
            Write-Host "$($u) : $($p)"
        }
    }
    if ((Get-Job -State Failed).count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "PCs with failed jobs:" -ForegroundColor Yellow
        (Get-Job -State Failed).Location | ForEach-Object {
            $u = userLookup -pcname $_
            $p = $_.ToString()
            Write-Host "$($u) : $($p)"
        }
    }
    if ($offlinePCs.Offline.Count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "Offline PCs:" -ForegroundColor Yellow
        $offlinePCs | ForEach-Object {
            $p = $_.Offline
            $u = userLookup -pcname $p
            Write-Host "$($u) : $($p)" 
        }
    }
    pause
}
Get-Job | Remove-Job # Cleanup jobs
$computers | Select-Object Computer, User, currentUsers, lastLogin, lastBoot | Out-GridView