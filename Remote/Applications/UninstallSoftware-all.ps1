#Execute
#################################################################################################
# Uninstall a remote program silently
#
# This script checks the registry for software to remove matching the search string
# it can remove both msi packages and run WinPE uninstallers.
# This uninstaller will automatically determine whether to run the uninstall.exe for the
# program to remove, or if it should remove the program with msiexec.
# A list of all installed programs can be created using the InstalledSoftware.ps1
# script. Make sure that the searches you pass to this script are unique enough to only remove
# what you want to remove.
#
# This uninstalls a program from ALL pcs. BE CAREFUL!
#
###############################################################################################
# Path to list of computer-names and user-names
$pcList = $env:COMPUTERSLIST
$adminName = $env:adminName
#################################################################################################
if ($cred) {} else { $cred = Get-Credential $adminName }
Write-Host "This script will remove a program from ALL computers" -ForegroundColor Red
$SoftToRemove = Read-Host "Software to Remove"
function userLookup ([string]$pcname){ return $userList.user[[array]::IndexOf($userList.computer, $pcname)] }

$unintallFunction = [ScriptBlock]::Create({
    function uninstall ($SoftToRemove = $null){
        if ($SoftToRemove -ne $null){
            $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { Get-ItemProperty $_.PSPath } | ? { ([String]$_).trim() -match ([String]$SoftToRemove).Trim() } | select UninstallString, DisplayName
            $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { Get-ItemProperty $_.PSPath } | ? { ([String]$_).trim() -match ([String]$SoftToRemove).Trim() } | select UninstallString, DisplayName
            $uninstall32c = @()
            $uninstall64c = @()
            $uninstall32 | ForEach-Object {
                if ( $_.UninstallString ){
                    $uninstall32c += $_.UninstallString
                }
            }
            $uninstall64 | ForEach-Object {
                if ( $_.UninstallString ){
                    $uninstall64c += $_.UninstallString
                }
            }
            function thisUninstall ($uninstArr) {
                $uninstArr | foreach {
                    $thisUninstall = $_ -Replace "msiexec.exe","" -replace "/I", "" -replace "/X", ""
                    $thisUninstall = $thisUninstall.Trim()
                    if ($thisUninstall -match ".exe"){
                        echo "$thisUninstall -arg /S"
                        Start-Process $thisUninstall -arg "/S" -Wait
                    } else {
                        echo "msiexec.exe -arg /X $thisUninstall /qn /norestart"
                        Start-Process "msiexec.exe" -arg "/X `"$thisUninstall`" /qn /norestart" -Wait
                    }
                }
            }
            thisUninstall $uninstall32c
            thisUninstall $uninstall64c
        }
    }
})

$uninstallBlock = [ScriptBlock]::Create($unintallFunction.ToString() + "uninstall `"$SoftToRemove`"") # Add quotations around user-entered string

$user = @()
$computer = @()
$computers = @()
$offlinePCs = @()
#import list of usernames and computernames from .csv
$userList = Import-CSV $pcList
$userList | ForEach-Object { $user += $_.user ; $computer += $_.computer }

for ( $i=0; $i -le ($user.Length - 1); $i++) {
	$iuser = $user[$i]
	$icomputer = $computer[$i]
	Write-Host "Attempting to uninstall `'$SoftToRemove`' from $icomputer ($iuser`'s PC)"
    if (!(Test-Connection -ComputerName $icomputer -BufferSize 16 -Count 1 -ErrorAction 0 -Quiet)) { 
        $offlinePC = New-Object PSObject -Property @{'Offline'=$icomputer}
        $offlinePCs += $offlinePC
        #Clear
    } else {
        if ($icomputer -ne $env:COMPUTERNAME){
            Invoke-Command -ComputerName $icomputer -Credential $cred -ScriptBlock $uninstallBLock -AsJob
        }
    }
    #Clear
}

Start-Sleep -Seconds 5
## Wait for jobs to finish
$loop = $true
$count = 0
$njobs = (Get-Job -State Running).count
Clear
Write-Host -NoNewLine "Waiting for all uninstall jobs ($njobs) to complete" -ForegroundColor White
While ($loop){
    if ((Get-Job -State Running).count -eq 0){ $loop = $false }
    if ($count -eq 50){ $loop = $false }
    $jobs = (Get-Job -State Running).count
    if ($jobs -ne $njobs){
        Clear
        Write-Host -NoNewLine "Waiting for all uninstall jobs ($jobs) to complete" -ForegroundColor White
        for ($i = 0 ; $i -lt $count + 1 ; $i ++){ Write-Host -NoNewline "." -ForegroundColor White }
        $njobs = $jobs
    } else {
        Write-Host -NoNewLine "." -ForegroundColor White
    }
    Start-Sleep -Seconds 1 # Wait for any running jobs to finish
    $count = $count + 1
}
Clear

Get-Job -State Completed | %{ $computers += (Get-Job -Id $_.Id | Receive-Job -ErrorAction SilentlyContinue) }

if (((Get-Job -State Failed).count -gt 0) -or ($offline.Length -gt 0) -or ((Get-Job -State Running).count -gt 0) -or ($offlinePCs.Offline.Count -gt 0)){
    if ((Get-Job -State Running).count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "PCs with unfinnished business:" -ForegroundColor Yellow
        (Get-Job -State Running).Location | Foreach {
            $u = userLookup -pcname $_
            $p = $_
            echo "$u : $p"
        }
    }
    if ((Get-Job -State Failed).count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "PCs with failed jobs:" -ForegroundColor Yellow
        (Get-Job -State Failed).Location | Foreach {
            userLookup -pcname $_
            $p = $_.ToString()
            echo "$u : $p"
        }
    }
    if ($offlinePCs.Offline.Count -gt 0){
        Start-Sleep -Milliseconds 200
        Write-Host "Offline PCs" -ForegroundColor Yellow
        $offlinePCs | foreach {
            $p = $_.Offline
            $u = userLookup -pcname $p
            echo "$u : $p" 
        }
    }
    pause
}

Get-Job | Remove-Job

