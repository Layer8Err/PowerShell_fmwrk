#Execute
#################################################################################################
# Uninstall a program silently
#
# This script checks the registry for software to remove matching the search string
# it can remove both msi packages and run WinPE uninstallers.
# This uninstaller will automatically determine whether to run the uninstall.exe for the
# program to remove, or if it should remove the program with msiexec.
# A list of all installed programs can be created using the query-InstalledSoftware-search.ps1
# script. Make sure that the searches you pass to this script are unique enough to only remove
# what you want to remove.
#################################################################################################

$SoftToRemove = Read-Host "Software to Remove"

## Script function for removing software
$script = 'function uninstall ($SoftToRemove = $null){
    if ($SoftToRemove -ne $null){
        $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { Get-ItemProperty $_.PSPath } | ? { $_ -match $SoftToRemove } | select UninstallString, DisplayName
        $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { Get-ItemProperty $_.PSPath } | ? { $_ -match $SoftToRemove } | select UninstallString, DisplayName
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
                    Start-Process $thisUninstall -arg "/S" -Wait
                } else {
                    Start-Process "msiexec.exe" -arg "/X `"$thisUninstall`" /qn /norestart" -Wait
                }
            }
        }
        thisUninstall $uninstall32c
        thisUninstall $uninstall64c
    }
}
uninstall ' + $SoftToRemove

$uninstallBlock = [ScriptBlock]::Create($script)

Write-Host "Uninstalling $SoftToRemove.." -ForegroundColor Yellow
Invoke-Command -ScriptBlock $uninstallBlock

Pause