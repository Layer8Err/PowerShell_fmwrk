#Execute
##############################################################################################
# Check Windows 10 AppX
#
# This script looks for Windows 10 errors and attempts to restore 
# system health without running a Windows 10 refresh.
# You will need to supply your own Windows 10 iso
##############################################################################################
$win10ISOPath = "\\server\share\isos\Windows10.iso"
$fixFromISO = $false # change to $true if the $win10ISOPath is valid
##############################################################################################

Write-Host "Beginning Preliminary System Scan..." -ForegroundColor Yellow
sfc /scannow

### Options for attempting to fix a machine from an ISO
if ($fixFromISO){
    Write-Host "Mounting Windows 10 ISO from the network..." -ForegroundColor Yellow
    $mountResult = Mount-DiskImage -ImagePath $win10ISOPath -PassThru
    $driveLetter = ($mountResult | Get-Volume).DriveLetter
    $esdPath = $driveLetter + ":\sources\install.esd"
    Write-Host "Beginning Image Servicing cleanp..." -ForegroundColor Yellow
    DISM /Online /Cleanup-Image /RestoreHealth /Source:"$esdPath":1 /LimitAccess

    Write-Host "Dismounting Windows 10 ISO..." -ForegroundColor Yellow
    Dismount-DiskImage -ImagePath $win10ISOPath    
}
### End attempting to fix from ISO

Write-Host "Re-initializing Metro Apps..." -ForegroundColor Yellow
Get-AppxPackage -AllUsers | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}
Write-Host "Getting Windows Store Metro App..."
Get-AppxPackage *WindowsStore* -AllUsers | Add-AppxPackage
Write-Host "Resetting Windows Store cache..."
wsreset

Write-Host "Running post Image Servicing system scan..." -ForegroundColor Yellow
sfc /scannow

Write-Host "Please reboot the machine" -ForegroundColor Red
$reboot = Read-Host -Prompt "Reboot now? [y/N]"
if ($reboot.ToLower().Substring(0,1) -eq "y"){
    shutdown /t 000 /r /f
}