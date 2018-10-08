#Execute
###############################################################################################
# Run silent office 365 repair
# 1/09/2017 modified to work with Office 2016
###############################################################################################

$pcname = Read-Host 'PC-Name to run office 365 repair'
$repairType = Read-Host 'Full repair (Y/N)'
$adminName = $env:adminName

if ($cred) {} else { $cred = Get-Credential $adminName }

if ($repairType.ToLower() -eq 'y'){
    Write-Host "Running full Office 365 repair on $pcname" # This throws an error if a process isn't running
    $remoteCmd = 'kill -name OUTLOOK -Force
    kill -name *lync* -Force
    kill -name officeclicktorun -Force
    kill -name EXCEL -Force
    kill -name UCA -Force
    kill -name UCMapi -Force
    kill -name WINWORD -Force
	$clickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\"
	$clickToRunExe = $clickToRunPath + "officeclicktorun.exe"
	& $clickToRunExe scenario=Repair platform=x86 culture=en-us DisplayLevel=False RepairType=FullRepair'
} else {
    Write-Host "Running quick Office 365 repair on $pcname" # This Works
    $remoteCmd = '$clickToRunPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\"
	$clickToRunExe = $clickToRunPath + "officeclicktorun.exe"
	& $clickToRunExe scenario=Repair platform=x86 culture=en-us DisplayLevel=False RepairType=QuickRepair'
}

$scriptBlock = [ScriptBlock]::Create($remoteCmd)
Invoke-Command -Computer $pcname -Credential $cred -ScriptBlock $scriptBlock