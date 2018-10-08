#Execute
#################################################################################################
# Uninstall a remote program silently
#
# This script checks the registry for software to remove matching the search string
# it can remove both msi packages and run WinPE uninstallers.
# This uninstaller will automatically determine whether to run the uninstall.exe for the
# program to remove, or if it should remove the program with msiexec.
# A list of all installed programs can be created using the query-InstalledSoftware-search.ps1
# script. Make sure that the searches you pass to this script are unique enough to only remove
# what you want to remove.
#################################################################################################
$adminName = $env:adminName
#################################################################################################
$pcname = Read-Host "PC-Name to remove software from"
$SoftToRemove = Read-Host "Software to Remove"

$uninstallScr = [ScriptBlock]::Create({
    function uninstall ($SoftToRemove = $null){
        if ($SoftToRemove -ne $null){
            $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { 
                Get-ItemProperty $_.PSPath 
            } | Where-Object { 
                $_ -match $SoftToRemove 
            } | Select-Object UninstallString, DisplayName
            $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object { 
                Get-ItemProperty $_.PSPath 
            } | Where-Object { 
                $_ -match $SoftToRemove 
            } | Select-Object UninstallString, DisplayName
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
                $uninstArr | ForEach-Object {
                    $thisUninstall = $_ -Replace "msiexec.exe","" -replace "/I", "" -replace "/X", ""
                    $thisUninstall = $thisUninstall.Trim()
                    if ($thisUninstall -match ".exe"){
                        Write-Output "$thisUninstall -arg /S"
                        Start-Process $thisUninstall -arg "/S" -Wait
                    } else {
                        Write-Output "msiexec.exe -arg /X $thisUninstall /qn /norestart"
                        Start-Process "msiexec.exe" -arg "/X `"$thisUninstall`" /qn /norestart" -Wait
                    }
                }
            }
            thisUninstall $uninstall32c
            thisUninstall $uninstall64c
        }
    }
    uninstall
})

$uninstallBlock = [ScriptBlock]::Create($uninstallScr.ToString() + " " + $SoftToRemove) # Build Uninstall Block with program to remove

if ((Get-WmiObject Win32_ComputerSystem).Name -ne $pcname){
    if ($cred) {} else { $cred = Get-Credential $adminName }
    ## Execute the scriptblock on the remote PC
    Write-Host "Attempting to uninstall `'$SoftToRemove`' from $pcname. Please wait..." -ForegroundColor Yellow
    Invoke-Command -ComputerName $pcname -Credential $cred -ScriptBlock $uninstallBLock
} else {
    Write-Host "Uninstalling software from your own machine should really be done through the control panel..." -ForegroundColor Yellow
    Invoke-Command -ScriptBlock $uninstallBLock
}

PAUSE